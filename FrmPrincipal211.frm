VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.MDIForm FrmPrincipal 
   BackColor       =   &H00808000&
   Caption         =   "Lego 1.0"
   ClientHeight    =   5955
   ClientLeft      =   1725
   ClientTop       =   1845
   ClientWidth     =   6465
   Icon            =   "FrmPrincipal211.frx":0000
   LinkTopic       =   "MDIForm1"
   MousePointer    =   11  'Hourglass
   ScrollBars      =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Ferramentas 
      Align           =   3  'Align Left
      Height          =   12330
      Left            =   0
      Negotiate       =   -1  'True
      ScaleHeight     =   12270
      ScaleWidth      =   2715
      TabIndex        =   1
      Top             =   360
      Width           =   2775
      Begin VB.ListBox ListaProg 
         Appearance      =   0  'Flat
         Height          =   1005
         Left            =   4500
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   1980
         Visible         =   0   'False
         Width           =   1260
      End
      Begin VB.Frame F1 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   7575
         Left            =   -60
         TabIndex        =   5
         Top             =   -60
         Width           =   2775
         Begin VB.TextBox Txt 
            Appearance      =   0  'Flat
            BackColor       =   &H0080C0FF&
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   990
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   3330
            Visible         =   0   'False
            Width           =   1125
         End
         Begin VB.ListBox L 
            Appearance      =   0  'Flat
            Height          =   615
            ItemData        =   "FrmPrincipal211.frx":0442
            Left            =   870
            List            =   "FrmPrincipal211.frx":0444
            Sorted          =   -1  'True
            TabIndex        =   8
            Top             =   3030
            Visible         =   0   'False
            Width           =   1650
         End
         Begin VB.CommandButton Botao 
            Height          =   240
            Left            =   2310
            Picture         =   "FrmPrincipal211.frx":0446
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   2790
            Visible         =   0   'False
            Width           =   210
         End
         Begin VB.ComboBox CboObj 
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   0
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   3540
            Width           =   2760
         End
         Begin MSComctlLib.Toolbar T 
            Height          =   1080
            Left            =   60
            TabIndex        =   6
            Top             =   330
            Width           =   2670
            _ExtentX        =   4710
            _ExtentY        =   1905
            ButtonWidth     =   926
            ButtonHeight    =   900
            ImageList       =   "ImageList1"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   12
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "A"
                  Object.ToolTipText     =   "Seleção"
                  ImageIndex      =   1
                  Style           =   1
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "B"
                  Object.ToolTipText     =   "Botão"
                  ImageIndex      =   2
                  Style           =   1
                  Object.Width           =   1e-4
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "C"
                  Object.ToolTipText     =   "Frame"
                  ImageIndex      =   3
                  Style           =   1
               EndProperty
               BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "D"
                  Object.ToolTipText     =   "Imagem"
                  ImageIndex      =   4
                  Style           =   1
               EndProperty
               BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "E"
                  Object.ToolTipText     =   "Palavras"
                  ImageIndex      =   5
                  Style           =   1
               EndProperty
               BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "F"
                  Object.ToolTipText     =   "Check"
                  ImageIndex      =   6
                  Style           =   1
               EndProperty
               BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "G"
                  Object.ToolTipText     =   "Texto"
                  ImageIndex      =   7
                  Style           =   1
               EndProperty
               BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "H"
                  Object.ToolTipText     =   "ComboBox"
                  ImageIndex      =   8
                  Style           =   1
               EndProperty
               BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "J"
                  Object.ToolTipText     =   "Lista Box"
                  ImageIndex      =   9
                  Style           =   1
               EndProperty
               BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.Visible         =   0   'False
                  Key             =   "I"
                  Object.ToolTipText     =   "Banco de Dados"
                  ImageIndex      =   10
                  Style           =   1
               EndProperty
               BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.Visible         =   0   'False
                  Key             =   "L"
                  Object.ToolTipText     =   "Tabelas (RecordSets)"
                  ImageIndex      =   11
               EndProperty
               BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "M"
                  Object.ToolTipText     =   "Timer"
                  ImageIndex      =   12
               EndProperty
            EndProperty
            Enabled         =   0   'False
         End
         Begin MSComDlg.CommonDialog CommonDialog1 
            Left            =   1590
            Top             =   -180
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid 
            Height          =   3645
            Left            =   0
            TabIndex        =   14
            Top             =   3930
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   6429
            _Version        =   393216
            Cols            =   3
            FormatString    =   "|Nomes             |Valores               "
            _NumberOfBands  =   1
            _Band(0).Cols   =   3
         End
         Begin MSComctlLib.TreeView Prog 
            Height          =   1815
            Left            =   -30
            TabIndex        =   15
            Top             =   1710
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   3201
            _Version        =   393217
            Indentation     =   159
            LabelEdit       =   1
            LineStyle       =   1
            Sorted          =   -1  'True
            Style           =   7
            FullRowSelect   =   -1  'True
            SingleSel       =   -1  'True
            ImageList       =   "ImageList3"
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label1 
            BackColor       =   &H00808000&
            Caption         =   "Projetos"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   16
            Top             =   1440
            Width           =   2865
         End
         Begin VB.Label Label4 
            BackColor       =   &H00808000&
            Caption         =   "Propriedades"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   0
            TabIndex        =   12
            Top             =   2250
            Width           =   2835
         End
         Begin VB.Label Label3 
            BackColor       =   &H00808000&
            Caption         =   "Ferramentas"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   60
            TabIndex        =   11
            Top             =   60
            Width           =   2835
         End
      End
      Begin VB.Frame F2 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   5595
         Left            =   6390
         TabIndex        =   10
         Top             =   3060
         Visible         =   0   'False
         Width           =   2805
         Begin MSComctlLib.ImageList Img3 
            Left            =   1710
            Top             =   1470
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   6
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FrmPrincipal211.frx":0530
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FrmPrincipal211.frx":0984
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FrmPrincipal211.frx":0F20
                  Key             =   ""
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FrmPrincipal211.frx":123C
                  Key             =   ""
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FrmPrincipal211.frx":1E90
                  Key             =   ""
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FrmPrincipal211.frx":276C
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.ImageList ImageList3 
            Left            =   2310
            Top             =   630
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   5
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FrmPrincipal211.frx":2BC0
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FrmPrincipal211.frx":3A14
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FrmPrincipal211.frx":3E68
                  Key             =   ""
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FrmPrincipal211.frx":4CBC
                  Key             =   ""
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FrmPrincipal211.frx":5110
                  Key             =   ""
               EndProperty
            EndProperty
         End
      End
      Begin VB.Label lp 
         Caption         =   "Label1"
         Height          =   570
         Left            =   2190
         TabIndex        =   3
         Top             =   4650
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.Label LRun 
         Height          =   300
         Left            =   360
         TabIndex        =   2
         Top             =   2976
         Width           =   684
      End
   End
   Begin MSComDlg.CommonDialog ComRun 
      Left            =   6405
      Top             =   5730
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   6090
      Top             =   3870
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrincipal211.frx":59EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrincipal211.frx":5B00
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrincipal211.frx":5C14
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrincipal211.frx":8260
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrincipal211.frx":A8AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrincipal211.frx":AF80
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrincipal211.frx":B51C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrincipal211.frx":BDF8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog Com 
      Left            =   6330
      Top             =   3090
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8040
      Top             =   3330
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483624
      ImageWidth      =   28
      ImageHeight     =   28
      MaskColor       =   16776960
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrincipal211.frx":C6D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrincipal211.frx":D058
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrincipal211.frx":D9DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrincipal211.frx":E360
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrincipal211.frx":ECE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrincipal211.frx":F668
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrincipal211.frx":FFEC
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrincipal211.frx":10970
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrincipal211.frx":112F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrincipal211.frx":11C78
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrincipal211.frx":12554
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrincipal211.frx":12E30
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   21480
      _ExtentX        =   37888
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "A"
            Object.ToolTipText     =   "Nova Tela"
            ImageIndex      =   1
            Style           =   5
            Object.Width           =   1e-4
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "NP"
                  Text            =   "Novo Projeto       "
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Key             =   "NT"
                  Text            =   "Nova Tela"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "B"
            Object.ToolTipText     =   "Abrir"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "C"
            Object.ToolTipText     =   "Salvar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "D"
            ImageIndex      =   4
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "E"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "F"
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "G"
            Object.ToolTipText     =   "Executar"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "H"
            Object.ToolTipText     =   "Para a Execução"
            ImageIndex      =   6
         EndProperty
      EndProperty
   End
   Begin VB.Menu MenuArquivo 
      Caption         =   "&Arquivo"
      Begin VB.Menu menuNovo 
         Caption         =   "&Novo"
         Begin VB.Menu MenuProjetog 
            Caption         =   "Projeto"
         End
         Begin VB.Menu MenuBranco 
            Caption         =   "-"
            Index           =   0
         End
         Begin VB.Menu MenuTela 
            Caption         =   "Tela"
            Enabled         =   0   'False
            Shortcut        =   ^N
         End
      End
      Begin VB.Menu nadkf 
         Caption         =   "-"
      End
      Begin VB.Menu MenuAbrir 
         Caption         =   "&Abrir"
         Shortcut        =   ^A
      End
      Begin VB.Menu MenuBranco2 
         Caption         =   "-"
      End
      Begin VB.Menu MenuSalvar 
         Caption         =   "&Salvar"
         Enabled         =   0   'False
         Shortcut        =   ^S
      End
      Begin VB.Menu MenuSalComo 
         Caption         =   "Salvar &Como ..."
         Enabled         =   0   'False
         Shortcut        =   ^B
      End
      Begin VB.Menu meudddd 
         Caption         =   "-"
      End
      Begin VB.Menu MenuCompli 
         Caption         =   "&Compile                                      "
         Enabled         =   0   'False
         Shortcut        =   {F9}
      End
      Begin VB.Menu haskfd 
         Caption         =   "-"
      End
      Begin VB.Menu MenuSair 
         Caption         =   "&Sair"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu MenuExibir 
      Caption         =   "Exibir"
      Begin VB.Menu MenuProjeto 
         Caption         =   "&Projetos"
         Enabled         =   0   'False
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnusss 
         Caption         =   "-"
      End
      Begin VB.Menu MenuPropriedade 
         Caption         =   "Propriedade"
         Enabled         =   0   'False
         Shortcut        =   {F4}
      End
      Begin VB.Menu MenuBr 
         Caption         =   "-"
      End
      Begin VB.Menu MenuCod 
         Caption         =   "Codigos"
         Shortcut        =   {F7}
      End
      Begin VB.Menu menuSep 
         Caption         =   "-"
      End
      Begin VB.Menu MenuFerra 
         Caption         =   "Ferramentas"
         Checked         =   -1  'True
         Enabled         =   0   'False
         Shortcut        =   ^{F7}
      End
      Begin VB.Menu pop 
         Caption         =   "-"
      End
      Begin VB.Menu MenuEdito 
         Caption         =   "&Editor de Menu"
         Enabled         =   0   'False
         Shortcut        =   ^M
      End
      Begin VB.Menu menufsaf 
         Caption         =   "-"
      End
      Begin VB.Menu MenuLimpa 
         Caption         =   "Limpar Figura"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu MenuFormatar 
      Caption         =   "&Formatar"
      Begin VB.Menu MenuEnviarTraz 
         Caption         =   "&Enviar Objeto para Traz"
         Enabled         =   0   'False
         Shortcut        =   ^T
      End
      Begin VB.Menu ppp 
         Caption         =   "-"
      End
      Begin VB.Menu MenuEnviar 
         Caption         =   "&Enviar Objeto para Frente"
         Enabled         =   0   'False
         Shortcut        =   ^F
      End
   End
   Begin VB.Menu menuexecutar 
      Caption         =   "&Executar"
      Begin VB.Menu menuComplile 
         Caption         =   "&Executar          "
         Enabled         =   0   'False
         Shortcut        =   {F5}
      End
      Begin VB.Menu menuParaExecutar 
         Caption         =   "&Para de Executar"
         Enabled         =   0   'False
         Shortcut        =   %{BKSP}
      End
      Begin VB.Menu menfadfds 
         Caption         =   "-"
      End
      Begin VB.Menu menuopcoes 
         Caption         =   "&Instruções Iniciais"
         Enabled         =   0   'False
         Shortcut        =   ^O
      End
   End
   Begin VB.Menu MenuInserir 
      Caption         =   "Inserir"
      Begin VB.Menu MenuProcedimentoPublic 
         Caption         =   "Procedimentos Publicos"
      End
   End
End
Attribute VB_Name = "FrmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RowT As Long
Dim a As String
Dim i As Long
Dim Oi As Long
Dim Row As String
Dim X As Long
Dim CboPass As Boolean
Dim Req As Boolean
Dim Passa As Boolean
Public NomeDoArquivoASerSalvo As String

Private Sub Ferramentas_CloseClick()
Dim H As Long, W As Long


W = FrmPrincipal.ScaleWidth + Ferramentas.Width
If FrmPrincipal.WindowState <> 1 Then
    FrmPrincipal.Left = W - FrmPrincipal.Width
End If
If FrmPrincipal.WindowState <> 1 Then
    FrmPrincipal.Left = W - FrmPrincipal.Width
End If
Ferramentas.Visible = False
MenuFerra.Checked = False
End Sub

Private Sub Botao_Click()
'Grid.Enabled = False
If InStr(1, UCase(Grid.TextMatrix(Grid.Row, 1)), "COR") <> 0 Then
    Com.ShowColor
    If Com.Color <> 0 Then
        Grid.TextMatrix(Grid.Row, 2) = Com.Color
        MontaPropriedade
    End If
Else
    L.Top = Botao.Top + Botao.Height + 20
    L.Width = Grid.ColWidth(2)
    L.Left = Grid.CellLeft + Grid.Left
    L.Visible = True
    L.SetFocus
End If
End Sub

Private Sub Eventos_DblClick()
'Dim FTela As New FrmCodigo
FrmCodigo.TxtCod.Text = Eventos.SelectedItem.Tag
FrmCodigo.Status.Panels(4).Text = Eventos.SelectedItem.Text
FrmCodigo.Caption = "Codigo - " & Eventos.SelectedItem.FullPath
FrmCodigo.Index.Caption = Eventos.SelectedItem.Index
FrmCodigo.Show
End Sub

Private Sub Grid_Click()
Dim VT(1 To 30) As String, X As Long, Tipo As Boolean
RowT = Grid.Row
If Grid.Col = 2 Then
    
    VT(1) = "Nome"
    VT(2) = "Borda"
    VT(3) = "Objeto 3D"
    VT(4) = "Cor Fundo"
    VT(5) = "Cor da Letra"
    VT(6) = "Texto"
    VT(7) = "Mascara"
    VT(8) = "Texto"
    VT(9) = "Imagem"
    VT(10) = "Borda"
    VT(11) = "Comprimir"
    VT(12) = "Legenda"
    VT(13) = "Botao Max"
    VT(14) = "Botao Min"
    VT(15) = "Botao Fechar"
    VT(16) = "Ordem"
    VT(17) = "TamX"
    VT(18) = "TamY"
    VT(19) = "PoxX"
    VT(20) = "PoxY"
    VT(21) = "Fonte"
    VT(22) = "Tamanho"
    L.Visible = False
    Botao.Visible = False
    For X = 1 To 22
        If UCase(VT(X)) = UCase(Grid.TextMatrix(Grid.Row, 1)) Then
            Select Case X
                Case 1, 6, 12, 16, 17, 18, 19, 20, 21
'                    Txt.Width = Grid.CellWidth
'                    Txt.Height = Grid.CellHeight
'                    Txt.Top = Grid.CellTop + Grid.Top
'                    Txt.Left = Grid.CellLeft + Grid.Left
'                    Txt.Text = Grid.Text
'                    Txt.SelStart = 0
'                    Txt.SelLength = Len(Txt.Text)
'                    Txt.Visible = True
'                    Txt.ZOrder vbBringToFront
'                    Txt.SetFocus
                                
                Case Else ' 2, 3, 10, 11, 13, 14, 15
                    Botao.Top = Grid.CellTop + Grid.Top - 10
                    Botao.Left = Grid.CellLeft + Grid.Left + Grid.ColWidth(2) - Botao.Width - 10
                    Botao.Visible = True
                    Botao.ZOrder vbBringToFront
                    L.Clear
                    L.AddItem "Sim"
                    L.AddItem "Não"
                    Tipo = False
            End Select
            Exit For
        End If
    Next X
   
End If

End Sub

Private Sub Grid_DblClick()
Dim VT(1 To 30) As String, X As Long, Tipo As Boolean

If Grid.Col = 2 Then

    VT(1) = "Nome"
    VT(2) = "Borda"
    VT(3) = "Objeto 3D"
    VT(4) = "Cor Fundo"
    VT(5) = "Cor da Letra"
    VT(6) = "Texto"
    VT(7) = "Mascara"
    VT(8) = "Texto"
    VT(9) = "Imagem"
    VT(10) = "Borda"
    VT(11) = "Comprimir"
    VT(12) = "Legenda"
    VT(13) = "Botao Max"
    VT(14) = "Botao Min"
    VT(15) = "Botao Fechar"
    VT(16) = "Ordem"
    VT(17) = "TamX"
    VT(18) = "TamY"
    VT(19) = "PoxX"
    VT(20) = "PoxY"
    VT(21) = "Fonte"
    VT(22) = "Tamanho"
    VT(23) = "Tempo"
    For X = 1 To 23
        If UCase(VT(X)) = UCase(Grid.TextMatrix(Grid.Row, 1)) Then
            Select Case X
                Case 1, 4, 5, 6, 12, 16, 17, 18, 19, 20, 21, 23
                    Txt.Width = Grid.CellWidth - 10
                    Txt.Height = Grid.CellHeight - 10
                    Txt.Top = Grid.CellTop + Grid.Top
                    Txt.Left = Grid.CellLeft + Grid.Left
                    Txt.Text = Grid.Text
                    Txt.SelStart = 0
                    Txt.SelLength = Len(Txt.Text)
                    Txt.Visible = True
                    Txt.ZOrder vbBringToFront
                    Txt.SetFocus
                Case Else ' 2, 3, 10, 11, 13, 14, 15
                    Botao.Top = Grid.CellTop + Grid.Top - 10
                    Botao.Left = Grid.CellLeft + Grid.Left + Grid.ColWidth(2) - Botao.Width - 10
                    Botao.Visible = True
                    Botao.ZOrder vbBringToFront
                    L.Clear
                    L.AddItem "Sim"
                    L.AddItem "Não"
                    Tipo = False
            End Select
            Exit For
        End If
    Next X
   
End If

End Sub

Private Sub Grid_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
    Grid_Click
End If
End Sub


Private Sub L_DblClick()
Grid.TextMatrix(Grid.Row, Grid.Col) = L.Text
MontaPropriedade
Grid.SetFocus
End Sub

Private Sub L_LostFocus()
L.Visible = False
'Grid.Enabled = True
End Sub

Private Sub lp_Change()
SalProjetoMdb False, NomeRun
End Sub

Private Sub LRun_Change()
Abrir CRun.ComandoRun
menuComplile_Click
End Sub

Private Sub MDIForm_Activate()
On Error Resume Next
If Passa = True Then
    If CRun.OpenExe = 0 Then
        Me.Visible = False
    End If
    If CRun.OpenExe <> 0 Then
        Me.MousePointer = 0
        NoOpen = 0
        If CRun.OpenExe <> 2 Then
            FrmApresentacao.Show 1
        Else
            Abrir CRun.ComandoRun
        End If
        Passa = False
        If NoOpen = 1 Then MenuAbrir_Click
        If NoOpen = 0 Then Exit Sub
        If NoOpen = 2 Then
            Abilita True
            Novo
            NomeProg = Titulo_Pjt
            NomeProg = UCase(Left(NomeProg, 1)) + Right(NomeProg, Len(NomeProg) - 1)
            FrmPrincipal.Prog.Nodes.Add , tvwChild, "a1", NomeProg, 5
        End If
    End If
End If
End Sub

Private Sub MDIForm_Initialize()
If CRun.OpenExe = 0 Then
    Me.Visible = False
    FrmPrincipal.Visible = False
    FrmPrincipal.WindowState = 0
    FrmPrincipal.Top = 90000
    FrmPrincipal.Left = 90000
    Exit Sub
End If

End Sub

Private Sub MDIForm_Load()
PosicaoDoFrmPrincipal = 0
NBanco = 0
NomeDoArquivoASerSalvo = ""
Req = False
ReDim FrmTela(999) As New Form2
ReDim Bancos(9999) As DBase
ReDim Rs(999) As DRs
ContTela = 0
TelaAtual = -1
Req = True
Passa = True
MenuFerra.Checked = True
If CRun.OpenExe = 0 Then
    Me.Visible = False
End If
Grid.ColWidth(0) = 0
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
End
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
End
End Sub

Private Sub MenuAbrir_Click()
Abrir
End Sub

Private Sub MenuCod_Click()
'FrmCodigo.Visible = True
'FrmCodigo.TxtCod.SetFocus
End Sub

Private Sub MenuCompli_Click()
On Error Resume Next
FrmCompilar.Show 1
End Sub

Private Sub menuComplile_Click()
If Dir("C:\TmpVar.Win") <> "" Then
    Kill "C:\TmpVar.Win"
End If
If ContTela = 0 Then
    MsgBox "Não a Telas para ser Executadas", vbInformation, App.Title
    Exit Sub
End If
PosicaoDoFrmPrincipal = Me.WindowState
Me.WindowState = 1
ContRun = 0
Start Nome_Da_Tela
End Sub

Private Sub MenuEdito_Click()
FrmMenu.Show 1
End Sub

Private Sub MenuEnviar_Click()
NovoObj.ZOrder vbBringToFront
Dim X As Long
For X = 0 To 7
    FrmTela(TelaAtual).Im(X).ZOrder vbSendToBack
    FrmTela(TelaAtual).Im(X).ZOrder vbBringToFront
Next X

End Sub

Private Sub MenuEnviarTraz_Click()
NovoObj.ZOrder vbSendToBack
Dim X As Long
For X = 0 To 7
    FrmTela(TelaAtual).Im(X).ZOrder vbSendToBack
Next X
End Sub

Private Sub MenuFerra_Click()
Ferramentas.Visible = True
FrmPrincipal.Height = (FrmPrincipal.ScaleHeight / 2) + 100
FrmPrincipal.Width = 3360
FrmPrincipal.Left = FrmPrincipal.ScaleWidth - FrmPrincipal.Width
FrmPrincipal.Top = (FrmPrincipal.ScaleHeight - FrmPrincipal.Height)
FrmPrincipal.Height = (FrmPrincipal.ScaleHeight / 2) - 100
FrmPrincipal.Width = 3360
FrmPrincipal.Left = FrmPrincipal.ScaleWidth - FrmPrincipal.Width
FrmPrincipal.Top = 0
MenuFerra.Checked = True
End Sub

Private Sub MenuLimpa_Click()
NovoObj.Picture = LoadPicture("")
End Sub

Private Sub menuopcoes_Click()
FrmOpcoes.Show 1
'FrmCodigo.Visible = True
'If FrmCodigo.WindowState <> 1 Then
    'FrmCodigo.Height = 6000
    'FrmCodigo.Width = 6000
'End If
'FrmCodigo.Cbo.Caption = "Sistema"
'FrmCodigo.Env.Caption = "Inicialização"
'FrmCodigo.Cod.SetFocus
End Sub

Private Sub menuParaExecutar_Click()
StopTela
End Sub

Private Sub MenuProcedimentoPublic_Click()
On Error GoTo T
Dim a As String
Dim X As Long
Dim Comp As String
Dim FTela As New FrmCodigo
a = InputBox("Nome do Procedimento Publico ?", "Lego 1.1 ")
If Trim(a) <> "" Then
    FrmCodigo.Eventos.Nodes.Add "PROD", tvwChild, "PROD." + a, a, 6

    For X = 1 To FrmCodigo.Eventos.Nodes.Count
        If UCase("PROD." + a) = UCase(FrmCodigo.Eventos.Nodes(X).Key) Then
            FrmCodigo.Eventos.Nodes(X).Selected = True
            FTela.TxtCod.Text = FrmCodigo.Eventos.Nodes(X).Tag
            FTela.TxtCod.SelStart = Len(FTela.TxtCod.Text)
            FTela.TxtCod.SetFocus
            Exit Sub
        End If
    Next X
End If
T:
If Err Then
    MsgBox "Impossivel Continuar , Procedimento já existente ou invalido ! ! !", vbInformation, App.Title
End If
End Sub

Private Sub MenuProjeto_Click()
FrmPrincipal.Visible = True
FrmPrincipal.Prog.SetFocus
End Sub

Private Sub MenuProjetog_Click()
Me.MousePointer = 0
NoOpen = 0
FrmApresentacao.CmdAbrir.Visible = False
FrmApresentacao.Show 1
Passa = False
If NoOpen = 1 Then MenuAbrir_Click
If NoOpen = 0 Then Exit Sub
If NoOpen = 2 Then
    Abilita True
    Novo
End If
End Sub

Private Sub MenuPropriedade_Click()
FrmPrincipal.Visible = True
FrmPrincipal.Grid.SetFocus

End Sub

Private Sub MenuSair_Click()
End
End Sub

Private Sub MenuSalComo_Click()
SalProjetoMdb True
End Sub

Private Sub MenuSalvar_Click()
On Error Resume Next
SalProjetoMdb False
End Sub

Private Sub MenuTela_Click()
On Error Resume Next

a = ""

Inicio:

a = InputBox("Nome da Tela :", App.Title, a)
If a = "" Then
    Exit Sub
End If

If FrmPrincipal.Prog.Nodes.Count = 0 Then GoTo Passa

Dim X1 As Long

For X1 = 0 To FrmPrincipal.Prog.Nodes.Count - 1
    If UCase(a) = UCase(FrmPrincipal.Prog.Nodes(X1 + 1).Text) Then
        MsgBox "Impossivel Criar uma tela com este nome, Pois ela já existe ! ! !", vbCritical, App.Title
        GoTo Inicio
        Exit Sub
    End If
Next X1

Passa:


FrmPrincipal.Prog.Nodes.Add "a1", tvwChild, a, a, 1


With FrmCodigo.Eventos
    Dim nodX As Node
    
    .Nodes.Add "MOD", tvwChild, UCase(a), a, 1

    
    .Nodes.Add UCase(a), tvwChild, UCase(a + ".2"), "Ao Clicar 2 Vezes", 2
    .Nodes.Add UCase(a), tvwChild, UCase(a + ".1"), "Ao Clicar 1 Vezes", 2
    .Nodes.Add UCase(a), tvwChild, UCase(a + ".Ganhar"), "Ao Ganhar o focu", 2
    .Nodes.Add UCase(a), tvwChild, UCase(a + ".Perder"), "Ao Peder o focu", 2
    .Nodes.Add UCase(a), tvwChild, UCase(a + ".Red"), "Ao Redimecionar a tela", 2
    '.Nodes.Add UCase(A), tvwChild, UCase(A + ".Escrever"), "Ao Escrever", 2
    .Nodes.Add UCase(a), tvwChild, UCase(a + ".Fechar"), "Ao Fechar a Tela", 2
    .Nodes.Add UCase(a), tvwChild, UCase(a + ".Carregar"), "Ao Carregar a Tela", 2
    
End With

If ContTela = 0 Or ContTela = -1 Then
    Nome_Da_Tela = a
    ContTela = 0
End If

Max(ContTela) = True
Min(ContTela) = True
Fecha(ContTela) = True
FrmTela(ContTela).Width = 7305
FrmTela(ContTela).Nome.Caption = a
FrmTela(ContTela).Cont = ContTela
FrmTela(ContTela).Caption = "Tela " + Str(ContTela)
FrmTela(ContTela).Tag = a
FrmTela(ContTela).Show vbModeless
If ContTela = 0 Then
    TelaAtual = 0
End If
TelaAtual = ContTela
Set NovoObj = FrmTela(ContTela)

ContTela = ContTela + 1
End Sub

Private Sub T_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim X As Long
If TelaAtual = -1 Then
    Exit Sub
End If
For X = 1 To T.Buttons.Count
    T.Buttons(X).Value = tbrUnpressed
Next X
Button.Value = tbrPressed
If Button.Index = 1 Then
    FrmTela(TelaAtual).MousePointer = 0
    Exit Sub
End If
FrmTela(TelaAtual).MousePointer = 2
End Sub

Private Sub SalProjeto()
On Error Resume Next
Dim Arq As String, X As Byte, Diretorios As String, NameArq As String
Dim ArqAux As String, Xy As Long, Yx As Long, Ob As Object
Dim M As MenuItem, ArqCodigo As String

Com.FileName = ""
Com.Filter = "Projeto do Lego (*.Leg) |*.Leg|Todos os Arquivo (*.*)|*.*"
Com.ShowSave
If Com.FileName <> "" Then
    X = Len(Com.FileName)
    Do While X <> 0
        If Mid(Com.FileName, X, 1) = "\" Then
            Diretorios = Left(Com.FileName, X)
            Exit Do
        End If
        X = X - 1
    Loop
    NameArq = Right(Com.FileName, Len(Com.FileName) - X)
    ArqCodigo = Left(NameArq, Len(NameArq) - 4) + ".afs"
    
    If Dir(Com.FileName) <> "" Then
        If MsgBox("Arquivo Existente , Deseja Substituir ???", vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbYes Then
            Kill Com.FileName
'            Exit Sub
        End If
    End If
    Arq = Com.FileName
    Escreva "Inicialização", "Iniciar", Nome_Da_Tela, Arq
    Escreva "Inicialização", "Plataforma", Plataforma, Arq
    Escreva "Projeto", "Nome EXE", "Projeto.exe", Arq
    Escreva "Projeto", "Versão do Lego", "1.0", Arq
    Escreva "Projeto", "Nome do Projeto", NameArq, Arq
    Escreva "Projeto", "Titulo", Titulo_Pjt, Arq
    Escreva "Projeto", "Autor", Autor, Arq
    Escreva "Projeto", "Senha", CripSenha(Senha), Arq
    
    Escreva "Arquivos", "Dir", Diretorios, Arq
    
    For X = 0 To ContTela - 1
        Escreva "Arquivos", "Arq" + Format(Str(X), "000"), FrmTela(X).Tag + ".neo", Arq
        ArqAux = Diretorios + FrmTela(X).Tag + ".neo"
        Set Ob = FrmTela(X)
        Escreva Ob.Tag, "Nome", Ob.Tag, ArqAux
        Escreva Ob.Tag, "Tipo", Ob.Name, ArqAux
        Escreva Ob.Tag, "Texto", Ob.Text, ArqAux
        Escreva Ob.Tag, "Comprimir", Ob.Stretch, ArqAux
        Escreva Ob.Tag, "Figura", Ob.Picture, ArqAux
        Escreva Ob.Tag, "Borda", Ob.BorderStyle, ArqAux
        Escreva Ob.Tag, "Legenda", Ob.Caption, ArqAux
        Escreva Ob.Tag, "TAMX", Ob.Height, ArqAux
        Escreva Ob.Tag, "TAMY", Ob.Width, ArqAux
        Escreva Ob.Tag, "Cor de Fundo", Ob.BackColor, ArqAux
        Escreva Ob.Tag, "Nome", Ob.Tag, ArqAux
        Escreva Ob.Tag, "POXX", Ob.Top, ArqAux
        Escreva Ob.Tag, "POXY", Ob.Left, ArqAux
        Escreva Ob.Tag, "Tamanho", Ob.FontSize, ArqAux
        Escreva Ob.Tag, "Estilo", Ob.FontName, ArqAux
        Escreva Ob.Tag, "Cor de Letra", Ob.ForeColor, ArqAux
        Escreva Ob.Tag, "Local", Ob.ToolTipText, ArqAux
        Escreva Ob.Tag, "Ordem", Ob.TabIndex, ArqAux
        Escreva Ob.Tag, "3D", Ob.Appearance, ArqAux
        Escreva Ob.Tag, "Mascara", Ob.PasswordChar, ArqAux
        Escreva Ob.Tag, "MAX", Max(FrmTela(X).Cont.Caption), ArqAux
        Escreva Ob.Tag, "MIN", Min(FrmTela(X).Cont.Caption), ArqAux
        Escreva Ob.Tag, "Fechar", Fecha(FrmTela(X).Cont.Caption), ArqAux
        Xy = 0
        For Each Ob In FrmTela(X)
            If UCase(Ob.Name) = "XMENU" Or Ob.Name = "Im" Or Ob.Name = "Focus" Or Ob.Name = "S" Or Ob.Name = "Nome" Or Ob.Name = "Cont" Then
                GoTo proximo
            ElseIf Ob.Index = 0 Then
                GoTo proximo
            End If
            Escreva "Objetos", "Nome" + Trim(Str(Xy)), Ob.Tag, ArqAux
            Escreva Ob.Tag, "Tipo", Ob.Name, ArqAux
            Escreva Ob.Tag, "Texto", Ob.Text, ArqAux
            Escreva Ob.Tag, "Comprimir", Ob.Stretch, ArqAux
            Escreva Ob.Tag, "Figura", Ob.ToolTipText, ArqAux
            Escreva Ob.Tag, "Borda", Ob.BorderStyle, ArqAux
            Escreva Ob.Tag, "Legenda", Ob.Caption, ArqAux
            Escreva Ob.Tag, "TAMX", Ob.Height, ArqAux
            Escreva Ob.Tag, "TAMY", Ob.Width, ArqAux
            Escreva Ob.Tag, "Cor de Fundo", Ob.BackColor, ArqAux
            Escreva Ob.Tag, "Nome", Ob.Tag, ArqAux
            Escreva Ob.Tag, "POXX", Ob.Top, ArqAux
            Escreva Ob.Tag, "POXY", Ob.Left, ArqAux
            Escreva Ob.Tag, "Tamanho", Ob.FontSize, ArqAux
            Escreva Ob.Tag, "Estilo", Ob.FontName, ArqAux
            Escreva Ob.Tag, "Cor de Letra", Ob.ForeColor, ArqAux
            Escreva Ob.Tag, "Local", Ob.ToolTipText, ArqAux
            Escreva Ob.Tag, "Ordem", Ob.TabIndex, ArqAux
            Escreva Ob.Tag, "3D", Ob.Appearance, ArqAux
            Escreva Ob.Tag, "Mascara", Ob.PasswordChar, ArqAux
            Xy = Xy + 1

proximo:
                If UCase(Ob.Name) = "XMENU" Then
                    Escreva "Menu", "Fonte", FrmTela(X).xMenu.ItemsFont, ArqAux
                    Escreva "Menu", "Borda", FrmTela(X).xMenu.Style, ArqAux
                    Escreva "Menu", "Selecao", FrmTela(X).xMenu.HighLightStyle, ArqAux
                    
                    For Yx = 0 To FrmTela(X).xMenu.MenuTree.Count
                        Set M = FrmTela(X).xMenu.MenuTree(Yx)
                        Escreva "Menu", "Legenda" + Trim(Str(Yx)), M.Caption, ArqAux
                        Escreva "Menu", "Ident" + Trim(Str(Yx)), M.Ident, ArqAux
                        Escreva "Menu", "Chave" + Trim(Str(Yx)), M.Name, ArqAux
                        Escreva "Menu", "Root" + Trim(Str(Yx)), M.RootIndex, ArqAux
                    Next Yx
                End If
            Next

    Next X
    SalvarEventos1 Diretorios + ArqCodigo
End If
End Sub

Private Sub Abrir1()
On Error Resume Next
Dim Arq As String, X As Byte, Diretorios As String, NameArq As String
Dim ArqAux As String, Xy As Long, Yx As Long, Ob As Object
Dim M As MenuItem, ArqCodigo As String, Nometela As String
Dim Tipo As String, TipoX As String
Dim Hp As Long, a As String, Index As Long
Dim SenhaAux As String
Dim A1 As String, A2 As String

Com.FileName = ""
Com.Filter = "Projeto do Lego (*.Leg) |*.Leg|Todos os Arquivo (*.*)|*.*"
Com.ShowOpen

If Com.FileName <> "" Then
    X = Len(Com.FileName)
    Do While X <> 0
        If Mid(Com.FileName, X, 1) = "\" Then
            Diretorios = Left(Com.FileName, X)
            Exit Do
        End If
        X = X - 1
    Loop
    NameArq = Right(Com.FileName, Len(Com.FileName) - X)
    ArqCodigo = Left(NameArq, Len(NameArq) - 4) + ".afs"
    
    If Dir(Com.FileName) = "" Then
        MsgBox "Arquivo Invalido", vbQuestion
        Exit Sub
    End If
    Arq = Com.FileName
    Nome_Da_Tela = Ler("Inicialização", "Iniciar", "", Arq)
    Plataforma = Ler("Inicialização", "Plataforma", "", Arq)
    'Escreva "Projeto", "Nome EXE", "Projeto.exe", Arq
    'Escreva "Projeto", "Versão do Lego", "1.0", Arq
    'Escreva "Projeto", "Nome do Projeto", NameArq, Arq
    SenhaAux = Senha
    A1 = Ler("Projeto", "Titulo", "", Arq)
    A2 = Ler("Projeto", "Autor", "", Arq)
    Senha = DescpSenha(Ler("Projeto", "Senha", "", Arq))
    If Senha <> "" Then
        SenhaOk = False
        FrmSenha.Auto.Caption = "Autor : " + A2
        FrmSenha.Caption = "Senha : " + A1
        FrmSenha.Show 1
        If SenhaOk = False Then
            Senha = SenhaAux
            Exit Sub
        End If
    End If
    Titulo_Pjt = Ler("Projeto", "Titulo", "", Arq)
    Autor = Ler("Projeto", "Autor", "", Arq)
   
    Diretorios = Ler("Arquivos", "Dir", App.Path, Arq)
    X = 0
    ContTela = 0
    X = 0
    FrmCodigo.Eventos.Refresh
    Do While FrmCodigo.Eventos.Nodes.Count <> 0
        FrmCodigo.Eventos.Nodes.Remove 1
        X = X + 1
    Loop
    FrmCodigo.Eventos.Nodes.Add , , "MOD", "Modulos", 4
    Eventos.Nodes.Add , , "PROD", "Procedimentos", 6
    X = 0
    FechaGeral = False
    Do While FrmPrincipal.Prog.Nodes.Count <> 0
        Unload FrmTela(X)
        FrmPrincipal.Prog.Nodes.Remove 1
        X = X + 1
    Loop
    FechaGeral = True
    X = 0
    Do While True
    
proximo:
        Nometela = Ler("Arquivos", "Arq" + Format(Str(X), "000"), ";=;", Arq)
        If Nometela = ";=;" Then
            Exit Do
        End If
            
        ArqAux = Diretorios + Nometela
        If Dir(ArqAux) = "" Then
            MsgBox "Impossivel Localizar M Arquivo " + ArqAux, vbCritical, App.Title
            X = X + 1
            GoTo proximo
        End If
        Err.Number = 0
        Load FrmTela(X)
        FrmTela(X).Tag = Left(Nometela, Len(Nometela) - 4)
        FrmTela(X).Visible = False
        FrmTela(X).Cont.Caption = X
        Set Ob = FrmTela(X)
        Ob.Tag = Ler(Ob.Tag, "Nome", "", ArqAux)
        Ob.BorderStyle = Ler(Ob.Tag, "Borda", "", ArqAux)
        Ob.Caption = Ler(Ob.Tag, "Legenda", Ob.Caption, ArqAux)
        Ob.Height = Ler(Ob.Tag, "TAMX", Ob.Height, ArqAux)
        Ob.Width = Ler(Ob.Tag, "TAMY", Ob.Width, ArqAux)
        Ob.BackColor = Ler(Ob.Tag, "Cor de Fundo", Ob.BackColor, ArqAux)
        Ob.Top = Ler(Ob.Tag, "POXX", Ob.Top, ArqAux)
        Ob.Left = Ler(Ob.Tag, "POXY", Ob.Left, ArqAux)
        Ob.FontSize = Ler(Ob.Tag, "Tamanho", Ob.FontSize, ArqAux)
        Ob.FontName = Ler(Ob.Tag, "Estilo", Ob.FontName, ArqAux)
        Ob.ForeColor = Ler(Ob.Tag, "Cor de Letra", Ob.ForeColor, ArqAux)
        Ob.Appearance = Ler(Ob.Tag, "3D", "1", ArqAux)
        Ob.Visible = True
        Max(X) = Ler(Ob.Tag, "MAX", "FALSE", ArqAux)
        Min(X) = Ler(Ob.Tag, "MIN", "FALSE", ArqAux)
        Fecha(X) = Ler(Ob.Tag, "Fechar", "FALSE", ArqAux)
        
        a = Ob.Tag
        FrmPrincipal.Prog.Nodes.Add , , a, a, 1


        With FrmCodigo.Eventos
            Dim nodX As Node
            
            .Nodes.Add "MOD", tvwChild, UCase(a), a, 1
        
            
            .Nodes.Add UCase(a), tvwChild, UCase(a + ".2"), "Ao Clicar 2 Vezes", 2
            .Nodes.Add UCase(a), tvwChild, UCase(a + ".1"), "Ao Clicar 1 Vezes", 2
            .Nodes.Add UCase(a), tvwChild, UCase(a + ".Ganhar"), "Ao Ganhar M focu", 2
            .Nodes.Add UCase(a), tvwChild, UCase(a + ".Perder"), "Ao Peder M focu", 2
            .Nodes.Add UCase(a), tvwChild, UCase(a + ".Red"), "Ao Redimecionar a tela", 2
            .Nodes.Add UCase(a), tvwChild, UCase(a + ".Escrever"), "Ao Escrever", 2
            .Nodes.Add UCase(a), tvwChild, UCase(a + ".Fechar"), "Ao Fechar a Tela", 2
            .Nodes.Add UCase(a), tvwChild, UCase(a + ".Carregar"), "Ao Carregar a Tela", 2
            
        End With
        ContTela = ContTela + 1

        Xy = 0
        For Hp = 0 To 9999
            TipoX = Ler("Objetos", "Nome" + Trim(Str(Hp)), ";=;", ArqAux)
            If TipoX = ";=;" Then
                Exit For
            End If
            Index = -1
            Tipo = Ler(TipoX, "Tipo", "", ArqAux)
            If Tipo = "" Then Exit For
   
            If Tipo = "Cmd" Then
                Index = 0
                Load FrmTela(X).Cmd(FrmTela(X).Cmd.Count)
                Set Ob = FrmTela(X).Cmd(FrmTela(X).Cmd.Count - 1)
            ElseIf Tipo = "Fm" Then
                Index = 1
                Load FrmTela(X).Fm(FrmTela(X).Fm.Count)
                Set Ob = FrmTela(X).Fm(FrmTela(X).Fm.Count - 1)
            ElseIf Tipo = "Img" Then
                Index = 2
                Load FrmTela(X).Img(FrmTela(X).Img.Count)
                Set Ob = FrmTela(X).Img(FrmTela(X).Img.Count - 1)
                Ob.Picture = Ob.Picture
            ElseIf Tipo = "Lbl" Then
                Index = 3
                Load FrmTela(X).Lbl(FrmTela(X).Lbl.Count)
                Set Ob = FrmTela(X).Lbl(FrmTela(X).Lbl.Count - 1)
            ElseIf Tipo = "Chk" Then
                Index = 4
                Load FrmTela(X).Chk(FrmTela(X).Chk.Count)
                Set Ob = FrmTela(X).Chk(FrmTela(X).Chk.Count - 1)
            ElseIf Tipo = "Cbo" Then
                Index = 5
                Load FrmTela(X).Cbo(FrmTela(X).Cbo.Count)
                Set Ob = FrmTela(X).Cbo(FrmTela(X).Cbo.Count - 1)
            ElseIf Tipo = "Txt" Then
                Index = 6
                Load FrmTela(X).Txt(FrmTela(X).Txt.Count)
                Set Ob = FrmTela(X).Txt(FrmTela(X).Txt.Count - 1)
            ElseIf Tipo = "Lst" Then
                Index = 7
                Load FrmTela(X).Lst(FrmTela(X).Lst.Count)
                Set Ob = FrmTela(X).Lst(FrmTela(X).Lst.Count - 1)
            End If
            
                
            Ob.Tag = TipoX
            Ob.Text = Ler(Ob.Tag, "Texto", "", ArqAux)
            Ob.Stretch = Ler(Ob.Tag, "Comprimir", Ob.Stretch, ArqAux)
            Ob.Picture = LoadPicture(Ler(Ob.Tag, "Figura", Ob.Picture, ArqAux))
            Ob.BorderStyle = Ler(Ob.Tag, "Borda", Ob.BorderStyle, ArqAux)
            Ob.Caption = Ler(Ob.Tag, "Legenda", Ob.Caption, ArqAux)
            Ob.Height = Ler(Ob.Tag, "TAMX", Ob.Height, ArqAux)
            Ob.Width = Ler(Ob.Tag, "TAMY", Ob.Width, ArqAux)
            Ob.BackColor = Ler(Ob.Tag, "Cor de Fundo", Ob.BackColor, ArqAux)
            Ob.Top = Ler(Ob.Tag, "POXX", Ob.Top, ArqAux)
            Ob.Left = Ler(Ob.Tag, "POXY", Ob.Left, ArqAux)
            Ob.FontSize = Ler(Ob.Tag, "Tamanho", Ob.FontSize, ArqAux)
            Ob.FontName = Ler(Ob.Tag, "Estilo", Ob.FontName, ArqAux)
            Ob.ForeColor = Ler(Ob.Tag, "Cor de Letra", Ob.ForeColor, ArqAux)
            Ob.ToolTipText = Ler(Ob.Tag, "Local", Ob.ToolTipText, ArqAux)
            Ob.TabIndex = Ler(Ob.Tag, "Ordem", Ob.TabIndex, ArqAux)
            Ob.Appearance = Ler(Ob.Tag, "3D", Ob.Appearance, ArqAux)
            Ob.PasswordChar = Ler(Ob.Tag, "Mascara", Ob.PasswordChar, ArqAux)
            Ob.Visible = True
            FrmCodigo.Eventos.Nodes.Add UCase(FrmTela(X).Tag), tvwChild, UCase(FrmTela(X).Tag + "." + Ob.Tag), Ob.Tag, 3

            With FrmCodigo.Eventos
            
                .Nodes.Add UCase(FrmTela(X).Tag + "." + Ob.Tag), tvwChild, UCase(FrmTela(X).Tag + "." + Ob.Tag + ".1"), "Ao Clicar 1 Vezes", 2
                If Index = 7 Or Index = 6 Or Index = 5 Or Index = 3 Or Index = 2 Or Index = 1 Then
                    .Nodes.Add UCase(FrmTela(X).Tag + "." + Ob.Tag), tvwChild, UCase(FrmTela(X).Tag + "." + Ob.Tag + ".2"), "Ao Clicar 2 Vezes", 2
                End If
                .Nodes.Add UCase(FrmTela(X).Tag + "." + Ob.Tag), tvwChild, UCase(FrmTela(X).Tag + "." + Ob.Tag + ".Ganhar"), "Ao Ganhar M focu", 2
                .Nodes.Add UCase(FrmTela(X).Tag + "." + Ob.Tag), tvwChild, UCase(FrmTela(X).Tag + "." + Ob.Tag + ".Perder"), "Ao Perder M focu", 2
                If Index = 6 Or 3 Then
                    .Nodes.Add UCase(FrmTela(X).Tag + "." + Ob.Tag), tvwChild, UCase(FrmTela(X).Tag + "." + Ob.Tag + ".Escrever"), "Ao Escrever", 2
                End If
                
            End With
            
            Xy = Xy + 1
        Next Hp
        FrmTela(X).xMenu.ItemsFont = Ler("Menu", "Fonte", FrmTela(X).xMenu.ItemsFont, ArqAux)
        FrmTela(X).xMenu.Style = Ler("Menu", "Borda", FrmTela(X).xMenu.Style, ArqAux)
        FrmTela(X).xMenu.HighLightStyle = Ler("Menu", "Selecao", FrmTela(X).xMenu.HighLightStyle, ArqAux)
        FrmTela(X).Visible = True
        FrmTela(X).Show
        FrmTela(X).Refresh
        For Yx = 1 To 9999
            Tipo = Ler("Menu", "Legenda" + Trim(Str(Yx)), ";=;", ArqAux)
            If Tipo = ";=;" Then Exit For
            Set M = New MenuItem
            M.Caption = Tipo
            M.Ident = Ler("Menu", "Ident" + Trim(Str(Yx)), M.Ident, ArqAux)
            M.Name = Ler("Menu", "Chave" + Trim(Str(Yx)), M.Name, ArqAux)
            M.RootIndex = Ler("Menu", "Root" + Trim(Str(Yx)), M.RootIndex, ArqAux)
            Nome = M.Caption
            InX = M.Ident
            M.Ident = InX
            M.Caption = Nome
            M.Name = Nome
            M.Accelerator = Nome
            M.Description = Nome
            If M.Ident = 0 Then
                FrmCodigo.Eventos.Nodes.Add UCase(FrmTela(TelaAtual).Tag), tvwChild, UCase(FrmTela(TelaAtual).Tag + "." + M.Caption + Trim(Str(M.Ident))), M.Caption, 5
            ElseIf M.Ident = 1 Then
                FrmCodigo.Eventos.Nodes.Add NomeMenu1, tvwChild, UCase(FrmTela(TelaAtual).Tag + "." + M.Caption + Trim(Str(M.Ident))), M.Caption, 5
            ElseIf M.Ident = 2 Then
                FrmCodigo.Eventos.Nodes.Add NomeMenu2, tvwChild, UCase(FrmTela(TelaAtual).Tag + "." + M.Caption + Trim(Str(M.Ident))), M.Caption, 5
            ElseIf M.Ident = 3 Then
                FrmCodigo.Eventos.Nodes.Add NomeMenu3, tvwChild, UCase(FrmTela(TelaAtual).Tag + "." + M.Caption + Trim(Str(M.Ident))), M.Caption, 5
            ElseIf M.Ident = 4 Then
                FrmCodigo.Eventos.Nodes.Add NomeMenu4, tvwChild, UCase(FrmTela(TelaAtual).Tag + "." + M.Caption + Trim(Str(M.Ident))), M.Caption, 5
            ElseIf M.Ident = 5 Then
                FrmCodigo.Eventos.Nodes.Add NomeMenu5, tvwChild, UCase(FrmTela(TelaAtual).Tag + "." + M.Caption + Trim(Str(M.Ident))), M.Caption, 5
            ElseIf M.Ident = 6 Then
                FrmCodigo.Eventos.Nodes.Add NomeMenu6, tvwChild, UCase(FrmTela(TelaAtual).Tag + "." + M.Caption + Trim(Str(M.Ident))), M.Caption, 5
            End If
            If M.Ident = 0 Then
                NomeMenu1 = UCase(FrmTela(TelaAtual).Tag + "." + M.Caption + Trim(Str(M.Ident)))
            ElseIf M.Ident = 1 Then
                NomeMenu2 = UCase(FrmTela(TelaAtual).Tag + "." + M.Name + Trim(Str(M.Ident)))
            ElseIf M.Ident = 2 Then
                NomeMenu3 = UCase(FrmTela(TelaAtual).Tag + "." + M.Name + Trim(Str(M.Ident)))
            ElseIf M.Ident = 3 Then
                NomeMenu4 = UCase(FrmTela(TelaAtual).Tag + "." + M.Name + Trim(Str(M.Ident)))
            ElseIf M.Ident = 4 Then
                NomeMenu5 = UCase(FrmTela(TelaAtual).Tag + "." + M.Name + Trim(Str(M.Ident)))
            ElseIf M.Ident = 5 Then
                NomeMenu6 = UCase(FrmTela(TelaAtual).Tag + "." + M.Name + Trim(Str(M.Ident)))
            End If
            FrmTela(X).xMenu.MenuTree.Add M
        Next Yx
        FrmTela(X).xMenu.Refresh
        FrmTela(X).SetFocus
        X = X + 1
    Loop
    BuscaEventosArq Diretorios + ArqCodigo
    Abilita True
End If

End Sub

Private Function CripSenha(Senha As String)
On Error Resume Next
Dim NovaSenha As String
Dim X As Long

NovaSenha = ""
If Senha = "" Then
    CripSenha = ""
    Exit Function
End If
For X = 1 To Len(Senha) + 1
    NovaSenha = NovaSenha + Chr(Asc(Mid(Senha, X, 1)) + 37)
Next X
CripSenha = NovaSenha
End Function
Private Function DescpSenha(Senha As String)
On Error Resume Next
Dim NovaSenha As String
Dim X As Long

NovaSenha = ""
If Senha = "" Then
    DescpSenha = ""
    Exit Function
End If
For X = 1 To Len(Senha)
    NovaSenha = NovaSenha + Chr(Asc(Mid(Senha, X, 1)) - 37)
Next X
DescpSenha = NovaSenha
End Function

Private Sub SalProjetoMdb(Tipo As Boolean, Optional sFile As String)
On Error Resume Next
Dim Arq As String, X As Byte, Diretorios As String, NameArq As String
Dim ArqAux As String, Xy As Long, Yx As Long, Ob As Object
Dim M As MenuItem, ArqCodigo As String
Dim Rs As Recordset
Inicio:
Com.FileName = ""
Com.Filter = "Projeto do Lego (*.Afs) |*.afs|Todos os Arquivo (*.*)|*.*"
If Tipo = True Then
    Com.ShowSave
Else
    If sFile <> "" Then
        Com.FileName = sFile
    Else
        If NomeDoArquivoASerSalvo = "" Then
            Com.ShowSave
            Tipo = True
        Else
            Com.FileName = NomeDoArquivoASerSalvo
        End If
    End If
End If
If Com.FileName <> "" Then
    'X = Len(Com.FileName)
'    Do While X <> 0
   '     If Mid(Com.FileName, X, 1) = "\" Then
   '         Diretorios = Left(Com.FileName, X)
   '         Exit Do
   '     End If
    '    X = X - 1
    'Loop
    NameArq = Com.FileName
    
    If Dir(Com.FileName) <> "" Then
        If Tipo = True Then
            If MsgBox("Arquivo Existente , Deseja Substituir ???", vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbYes Then
                Kill Com.FileName
    '            Exit Sub
            Else
                GoTo Inicio
            End If
        Else
            Kill Com.FileName
        End If
    End If
    Screen.MousePointer = 11
    Arq = Com.FileName
    CriarBanco Arq
    Set Rs = Banco.OpenRecordset("Config")
    
    Rs.AddNew
    Rs!Banco = LocalBancodeDados
    Rs!Iniciar = Nome_Da_Tela
    Rs!Plataforma = Plataforma
    Rs!Exe = "Projeto.exe"
    Rs!Titulo = Titulo_Pjt
    Rs!Autor = Autor
    Rs!Senha = CripSenha(Senha)
    Rs!Esc = TeclaEsc
    Rs!Enter = TeclaEnter
    Rs.Update
    If ContTela = -1 Then GoTo Fim
    For X = 0 To ContTela - 1
        Set Ob = FrmTela(X)
        CriarEstrutura FrmTela(X).Tag
        Set Rs = Banco.OpenRecordset(FrmTela(X).Tag)
        Rs.AddNew
        Rs!Nome = Ob.Tag
        Rs!Tipo = Ob.Name
        Rs!Borda = Ob.BorderStyle
        Rs!Legenda = Ob.Caption
        Rs.Fields("TamX") = Ob.Height
        Rs.Fields("TamY") = Ob.Width
        Rs.Fields("Cor Fundo") = Ob.BackColor
        Rs.Fields("PoxX") = Ob.Top
        Rs.Fields("PoxY") = Ob.Left
        Rs.Fields("Tamanho") = Val(Ob.FontSize)
        Rs.Fields("Estilo") = Ob.FontName
        Rs.Fields("Cor da Letra") = Ob.ForeColor
        Rs.Fields("3D") = Ob.Appearance
        Rs.Fields("BotaoMAX") = Max(FrmTela(X).Cont.Caption)
        Rs.Fields("BotaoMIN") = Min(FrmTela(X).Cont.Caption)
        Rs.Fields("BotaoFechar") = Fecha(FrmTela(X).Cont.Caption)
        Rs!Imagem = BuscaImg(Ob.Icon)
        Rs.Update
        Xy = 0
        For Each Ob In FrmTela(X)
            If UCase(Ob.Name) = "XMENU" Or Ob.Name = "Im" Or Ob.Name = "Focus" Or Ob.Name = "S" Or Ob.Name = "Nome" Or Ob.Name = "Cont" Then
                GoTo proximo
            ElseIf Ob.Index = 0 Then
                GoTo proximo
            End If
            Set Rs = Banco.OpenRecordset(FrmTela(X).Tag)
            Rs.AddNew
            Rs!Nome = Ob.Tag
            Rs!Tipo = Ob.Name
            Rs!Texto = Ob.Text
            If UCase(Ob.Name) = "BANCOIMG" Then
                Rs!Texto = Bancos(Ob.Index).Local
                Rs!Nome = Bancos(Ob.Index).Nome
            End If
            Rs!Comprimir = Ob.Stretch
            Rs!Imagem = BuscaImg(Ob.Picture)
            Rs!Borda = Ob.BorderStyle
            Rs!Legenda = Ob.Caption
            Rs.Fields("TamX") = Ob.Height
            Rs.Fields("TamY") = Ob.Width
            Rs.Fields("Cor Fundo") = Ob.BackColor
            Rs.Fields("PoxX") = Ob.Top
            Rs.Fields("PoxY") = Ob.Left
            Rs.Fields("Tamanho") = Ob.FontSize
            Rs.Fields("Estilo") = Ob.FontName
            Rs.Fields("Cor dd Letra") = Ob.ForeColor
            Rs.Fields("3D") = Ob.Appearance
            Rs.Fields("Ordem") = Ob.TabIndex
            Rs.Fields("Mascara") = Ob.PasswordChar
            Rs.Fields("Order") = Ob.TabIndex
            If UCase(Ob.Name) = "IMGRECORD" Then
                Rs!Texto = TabRec(Ob.Index).BancoDb
                Rs!Nome = TabRec(Ob.Index).Nome
                Rs!Imagem = TabRec(Ob.Index).Codicao
                Rs!Borda = TabRec(Ob.Index).Ordem
                Rs!Legenda = TabRec(Ob.Index).Tabela
            End If
                
                
            Rs.Update
            Xy = Xy + 1

proximo:
                If UCase(Ob.Name) = "XMENU" Then
                    Set Rs = Banco.OpenRecordset("Menu-" + FrmTela(X).Tag)
                    Rs.AddNew
                    Rs!Legenda = FrmTela(X).xMenu.ItemsFont
                    Rs!Chave = FrmTela(X).xMenu.Style
                    Rs!Root = FrmTela(X).xMenu.HighLightStyle
                    Rs.Update
                    For Yx = 1 To FrmTela(X).xMenu.MenuTree.Count
                        Set M = FrmTela(X).xMenu.MenuTree(Yx)
                        Rs.AddNew
                        Rs!Legenda = M.Caption
                        Rs!Ident = M.Ident
                        Rs!Chave = M.Name
                        Rs!Root = M.RootIndex
                        Rs.Update
                    Next Yx
                End If
            Next

    Next X
    SalvarEventos
Fim:
    Banco.Close
    Screen.MousePointer = 0
    NomeDoArquivoASerSalvo = Com.FileName
End If
Screen.MousePointer = 0
End Sub


Private Sub Novo()
Dim X As Long
X = 1
FechaGeral = False
Do While FrmCodigo.Eventos.Nodes.Count <> 0
    FrmCodigo.Eventos.Nodes.Remove 1
    X = X + 1
Loop
X = 1
Do While FrmPrincipal.Prog.Nodes.Count <> 0
    Unload FrmTela(X - 1)
    FrmPrincipal.Prog.Nodes.Remove 1
    X = X + 1
Loop

TeclaEnter = 0
TeclaEsc = 0
FechaGeral = True
ContTela = -1
TelaAtual = -1
FrmCodigo.Eventos.Nodes.Add , , "MOD", "Modulos", 4
FrmCodigo.Eventos.Nodes.Add , , "PROD", "Procedimentos", 6
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
If Button.Index = 1 Then
    If Toolbar1.Buttons(1).ButtonMenus(2).Enabled = True Then
        Call MenuProjetog_Click
    Else
        Call MenuTela_Click
        
    End If
ElseIf Button.Index = 2 Then
    Abrir
ElseIf Button.Index = 3 Then
    SalProjetoMdb False
ElseIf Button.Index = 7 Then
    menuComplile_Click
ElseIf Button.Index = 8 Then
    StopTela
End If
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
If ButtonMenu.Index = 1 Then
    Call MenuProjetog_Click
ElseIf ButtonMenu.Index = 2 Then
    Call MenuTela_Click
End If
End Sub
Private Sub Abrir(Optional sFile As String)

On Error Resume Next
Dim Arq As String, X As Byte, Diretorios As String, NameArq As String
Dim ArqAux As String, Xy As Long, Yx As Long, Ob As Object
Dim M As MenuItem, ArqCodigo As String, Nometela As String
Dim Tipo As String, TipoX As String
Dim Hp As Long, a As String, Index As Long
Dim SenhaAux As String
Dim A1 As String, A2 As String
Dim Rs As Recordset
Dim TabelaIndex As Long

Com.FileName = sFile
Com.Filter = "Projeto do Lego (*.Afs) |*.Afs|Todos os Arquivo (*.*)|*.*"
If sFile = "" Then
    Com.ShowOpen
End If
If CRun.OpenExe = 0 Then Me.Visible = False
If Com.FileName <> "" Then
    X = Len(Com.FileName)
    'Do While X <> 0
    '    If Mid(Com.FileName, X, 1) = "\" Then
    '        Diretorios = Left(Com.FileName, X)
    '        Exit Do
    '    End If
    '    X = X - 1
    'Loop
    NameArq = Com.FileName
    NomeDoArquivoASerSalvo = Com.FileName
    If Dir(Com.FileName) = "" Then
        MsgBox "Arquivo Invalido", vbQuestion
        Exit Sub
    End If
    Arq = Com.FileName
    Err.Number = 0
    
    'Abre o Arquivo
    Set Banco = OpenDatabase(Arq)
    If Err.Number <> 0 Then
        MsgBox "Arquivo em Formato Invalido ! ! !", vbCritical, App.Title
        Banco.Close
        Exit Sub
    End If
    
    'Abre as Configurações
    Set Rs = Banco.OpenRecordset("Config")
    If Rs.RecordCount = 0 Then
        MsgBox "Impossivel localizar Paramentos a do Sistema", vbCritical, App.Title
        Rs.Close
        Banco.Close
        Exit Sub
    End If
    
    SenhaAux = Senha
    A1 = Rs!Titulo
    A2 = Rs!Autor
    Senha = DescpSenha(Rs!Senha)
    If Senha <> "" Then
        SenhaOk = False
        FrmSenha.Auto.Caption = "Autor : " + A2
        FrmSenha.Caption = "Senha : " + A1
        Rs.Close
        Banco.Close
        
        FrmSenha.Show 1
        If SenhaOk = False Then
            Senha = SenhaAux
            Exit Sub
        End If
    End If
    Set Banco = OpenDatabase(Arq)
    Set Rs = Banco.OpenRecordset("Config")
    Nome_Da_Tela = Rs!Iniciar
    Plataforma = Rs!Plataforma
    Titulo_Pjt = Rs!Titulo
    Autor = Rs!Autor
    If CRun.OpenExe = 0 Then
        Me.Caption = Titulo_Pjt
    End If
    TeclaEnter = Rs!Enter
    TeclaEsc = Rs!Esc
    LocalBancodeDados = Rs!Banco
    X = 0
    ContTela = 0
    X = 0
    If CRun.OpenExe = 0 Then
        Me.Visible = False
    End If
    FrmCodigo.Eventos.Refresh
    If CRun.OpenExe = 0 Then
        Me.Visible = False
    End If
    Do While FrmCodigo.Eventos.Nodes.Count <> 0
        FrmCodigo.Eventos.Nodes.Remove 1
        X = X + 1
    Loop
    FrmCodigo.Eventos.Nodes.Add , , "MOD", "Modulos", 4
    FrmCodigo.Eventos.Nodes.Add , , "PROD", "Procedimentos", 6
    X = 0
    FechaGeral = False
    Do While FrmPrincipal.Prog.Nodes.Count <> 0
        Unload FrmTela(X)
        FrmPrincipal.Prog.Nodes.Remove FrmPrincipal.Prog.Nodes.Count
        X = X + 1
    Loop
    
    FechaGeral = True
    X = 0
    TabelaIndex = 0
    
    NomeProg = Titulo_Pjt
    NomeProg = UCase(Left(NomeProg, 1)) + Right(NomeProg, Len(NomeProg) - 1)
    FrmPrincipal.Prog.Nodes.Add , tvwChild, "a1", NomeProg, 5
    
    
    
    Do While TabelaIndex <> Banco.TableDefs.Count
proximo:
        If CRun.OpenExe = 0 Then FrmPrincipal.Visible = False
        Nometela = Banco.TableDefs(TabelaIndex).Name
        
        If Nometela = "Config" Or Nometela = "Lego" Or Nometela = "MSysACEs" Or Nometela = "MSysModules" Or Nometela = "MSysModules2" Or Nometela = "MSysObjects" Or Nometela = "MSysQueries" Or Nometela = "MSysRelationships" Or Left(Nometela, 5) = "Menu-" Or Nometela = "MSysAccessObjects" Then
            TabelaIndex = TabelaIndex + 1
            If TabelaIndex >= Banco.TableDefs.Count Then Exit Do
            GoTo proximo
        End If
            
        'Inicio da Criação
        
        Err.Number = 0
        Load FrmTela(X)
        If CRun.OpenExe = 0 Then Me.Visible = False
        FrmTela(X).Tag = Nometela
        FrmTela(X).Visible = False
        FrmTela(X).Cont.Caption = X
        Set Ob = FrmTela(X)
        Set Rs = Banco.OpenRecordset(Nometela)
        Ob.Appearance = Rs.Fields("3D")
        Ob.Tag = Nometela
        Ob.BorderStyle = Rs!Borda
        Ob.Caption = Rs!Legenda
        Ob.Height = Rs.Fields("TamX")
        Ob.Width = Rs.Fields("TamY")
        Ob.BackColor = Rs.Fields("Cor Fundo")
        Ob.Top = Rs.Fields("PoxX")
        Ob.Left = Rs.Fields("PoxY")
        Ob.FontSize = Rs.Fields("Tamanho")
        Ob.FontName = Rs.Fields("Estilo")
        Ob.ForeColor = Rs.Fields("Cor da Letra")
        AbreFig Rs!Imagem, Ob, False
        'Ob.Visible = True
        Max(X) = Rs.Fields("BotaoMAX")
        Min(X) = Rs.Fields("BotaoMIN")
        Fecha(X) = Rs.Fields("BotaoFechar")
        
        a = Ob.Tag
        FrmPrincipal.Prog.Nodes.Add "a1", tvwChild, a, a, 1


        With FrmCodigo.Eventos
            Dim nodX As Node
            
            .Nodes.Add "MOD", tvwChild, UCase(a), a, 1
        
            
            .Nodes.Add UCase(a), tvwChild, UCase(a + ".2"), "Ao Clicar 2 Vezes", 2
            .Nodes.Add UCase(a), tvwChild, UCase(a + ".1"), "Ao Clicar 1 Vezes", 2
            .Nodes.Add UCase(a), tvwChild, UCase(a + ".Ganhar"), "Ao Ganhar M focu", 2
            .Nodes.Add UCase(a), tvwChild, UCase(a + ".Perder"), "Ao Peder M focu", 2
            .Nodes.Add UCase(a), tvwChild, UCase(a + ".Red"), "Ao Redimecionar a tela", 2
            .Nodes.Add UCase(a), tvwChild, UCase(a + ".Escrever"), "Ao Escrever", 2
            .Nodes.Add UCase(a), tvwChild, UCase(a + ".Fechar"), "Ao Fechar a Tela", 2
            .Nodes.Add UCase(a), tvwChild, UCase(a + ".Carregar"), "Ao Carregar a Tela", 2
            
        End With
        ContTela = ContTela + 1

        Xy = 0
        Rs.MoveNext
        If Rs.EOF = True Then GoTo ProximaTela
        
        Do While Not Rs.EOF
            Index = -1
            Tipo = Rs!Tipo
  
            If Tipo = "Cmd" Then
                Index = 0
                Load FrmTela(X).Cmd(FrmTela(X).Cmd.Count)
                Set Ob = FrmTela(X).Cmd(FrmTela(X).Cmd.Count - 1)
            ElseIf Tipo = "Fm" Then
                Index = 1
                Load FrmTela(X).Fm(FrmTela(X).Fm.Count)
                Set Ob = FrmTela(X).Fm(FrmTela(X).Fm.Count - 1)
            ElseIf Tipo = "Img" Then
                Index = 2
                Load FrmTela(X).Img(FrmTela(X).Img.Count)
                Set Ob = FrmTela(X).Img(FrmTela(X).Img.Count - 1)
            ElseIf Tipo = "Lbl" Then
                Index = 3
                Load FrmTela(X).Lbl(FrmTela(X).Lbl.Count)
                Set Ob = FrmTela(X).Lbl(FrmTela(X).Lbl.Count - 1)
            ElseIf Tipo = "Chk" Then
                Index = 4
                Load FrmTela(X).Chk(FrmTela(X).Chk.Count)
                Set Ob = FrmTela(X).Chk(FrmTela(X).Chk.Count - 1)
            ElseIf Tipo = "Cbo" Then
                Index = 5
                Load FrmTela(X).Cbo(FrmTela(X).Cbo.Count)
                Set Ob = FrmTela(X).Cbo(FrmTela(X).Cbo.Count - 1)
            ElseIf Tipo = "Txt" Then
                Index = 6
                Load FrmTela(X).Txt(FrmTela(X).Txt.Count)
                Set Ob = FrmTela(X).Txt(FrmTela(X).Txt.Count - 1)
            ElseIf Tipo = "Lst" Then
                Index = 7
                Load FrmTela(X).Lst(FrmTela(X).Lst.Count)
                Set Ob = FrmTela(X).Lst(FrmTela(X).Lst.Count - 1)
            ElseIf Tipo = "BancoImg" Then
                Index = 8
                Load FrmTela(X).BancoImg(FrmTela(X).BancoImg.Count)
                Set Ob = FrmTela(X).BancoImg(FrmTela(X).BancoImg.Count - 1)
            ElseIf Tipo = "ImgRecord" Then
                Index = 9
                Load FrmTela(X).ImgRecord(FrmTela(X).ImgRecord.Count)
                Set Ob = FrmTela(X).ImgRecord(FrmTela(X).ImgRecord.Count - 1)
            End If
            FrmTela(X).O.Caption = Index
                
            Ob.Tag = Rs!Nome
            If Index = 9 Then
                GoTo p1
            End If
            Ob.Appearance = Rs.Fields("3D")
            Ob.Text = Rs!Texto
            Ob.Stretch = Rs!Comprimir
            AbreFig Rs!Imagem, Ob, True
            Ob.BorderStyle = Rs!Borda
            Ob.Caption = Rs!Legenda
            Ob.Height = Rs!TamX
            Ob.Width = Rs!TamY
            Ob.BackColor = Rs.Fields("Cor Fundo")
            Ob.Top = Rs!PoxX
            Ob.Left = Rs!PoxY
            Ob.FontSize = Rs!Tamanho
            Ob.FontName = Rs!Estilo
            Ob.ForeColor = Rs.Fields("Cor da Letra")
            Ob.TabIndex = Rs!Ordem
            Ob.PasswordChar = Rs!Mascara
            
            If Index = 8 Then
                Bancos(Ob.Index).Local = Rs!Texto
                Bancos(Ob.Index).Nome = Rs!Nome
            End If
            
            If Index = 9 Then
p1:
                TabRec(Ob.Index).BancoDb = Rs!Texto
                TabRec(Ob.Index).Nome = Rs!Nome
                TabRec(Ob.Index).Codicao = Rs!Imagem
                TabRec(Ob.Index).Ordem = Rs!Borda
                TabRec(Ob.Index).Tabela = Rs!Legenda
            End If
            Ob.Visible = True
            FrmCodigo.Eventos.Nodes.Add UCase(FrmTela(X).Tag), tvwChild, UCase(FrmTela(X).Tag + "." + Ob.Tag), Ob.Tag, 3

            With FrmCodigo.Eventos
                If Index <> 8 And Index <> 9 Then
                    .Nodes.Add UCase(FrmTela(X).Tag + "." + Ob.Tag), tvwChild, UCase(FrmTela(X).Tag + "." + Ob.Tag + ".1"), "Ao Clicar 1 Vezes", 2
                    If Index = 7 Or Index = 6 Or Index = 5 Or Index = 3 Or Index = 2 Or Index = 1 Then
                        .Nodes.Add UCase(FrmTela(X).Tag + "." + Ob.Tag), tvwChild, UCase(FrmTela(X).Tag + "." + Ob.Tag + ".2"), "Ao Clicar 2 Vezes", 2
                    End If
                    .Nodes.Add UCase(FrmTela(X).Tag + "." + Ob.Tag), tvwChild, UCase(FrmTela(X).Tag + "." + Ob.Tag + ".Ganhar"), "Ao Ganhar M focu", 2
                    .Nodes.Add UCase(FrmTela(X).Tag + "." + Ob.Tag), tvwChild, UCase(FrmTela(X).Tag + "." + Ob.Tag + ".Perder"), "Ao Perder M focu", 2
                    If Index = 6 Or 3 Then
                        .Nodes.Add UCase(FrmTela(X).Tag + "." + Ob.Tag), tvwChild, UCase(FrmTela(X).Tag + "." + Ob.Tag + ".Escrever"), "Ao Escrever", 2
                    End If
                End If
            End With
            
            Xy = Xy + 1
            Rs.MoveNext
        Loop
        Set Rs = Banco.OpenRecordset("Select * From " & FrmTela(X).Tag & "", dbOpenDynaset)
        Do While Not Rs.EOF
            For Each Ob In FrmTela(X)
                If UCase(Rs!Nome) = UCase(Ob.Tag) Then
                    Ob.TabIndex = Rs!Ordem
                    Exit For
                End If
            Next
            Rs.MoveNext
        Loop
            
            
ProximaTela:
        Set Rs = Banco.OpenRecordset("Menu-" + FrmTela(X).Tag)
        If Rs.EOF = True Then GoTo ProximaTela1:
        FrmTela(X).Visible = True
        FrmTela(X).xMenu.ItemsFont = Rs!Legenda
        FrmTela(X).xMenu.Style = Rs!Chave
        FrmTela(X).xMenu.HighLightStyle = Rs!Root

        FrmTela(X).Show
        If CRun.ComandoRun <> 0 Then
            FrmTela(X).Refresh
        End If
        Rs.MoveNext
        If Rs.EOF = True Then GoTo ProximaTela1
        
        Do While Not Rs.EOF
            
            Set M = New MenuItem
            M.Caption = Rs!Legenda
            M.Ident = Rs!Ident
            M.Name = Rs!Chave
            M.RootIndex = Rs!Root
            Nome = M.Caption
            InX = M.Ident
            M.Ident = InX
            M.Caption = Nome
            M.Name = Nome
            M.Accelerator = Nome
            M.Description = Nome
            If M.Ident = 0 Then
                FrmCodigo.Eventos.Nodes.Add UCase(FrmTela(TelaAtual).Tag), tvwChild, UCase(FrmTela(TelaAtual).Tag + "." + M.Caption + Trim(Str(M.Ident))), M.Caption, 5
            ElseIf M.Ident = 1 Then
                FrmCodigo.Eventos.Nodes.Add NomeMenu1, tvwChild, UCase(FrmTela(TelaAtual).Tag + "." + M.Caption + Trim(Str(M.Ident))), M.Caption, 5
            ElseIf M.Ident = 2 Then
                FrmCodigo.Eventos.Nodes.Add NomeMenu2, tvwChild, UCase(FrmTela(TelaAtual).Tag + "." + M.Caption + Trim(Str(M.Ident))), M.Caption, 5
            ElseIf M.Ident = 3 Then
                FrmCodigo.Eventos.Nodes.Add NomeMenu3, tvwChild, UCase(FrmTela(TelaAtual).Tag + "." + M.Caption + Trim(Str(M.Ident))), M.Caption, 5
            ElseIf M.Ident = 4 Then
                FrmCodigo.Eventos.Nodes.Add NomeMenu4, tvwChild, UCase(FrmTela(TelaAtual).Tag + "." + M.Caption + Trim(Str(M.Ident))), M.Caption, 5
            ElseIf M.Ident = 5 Then
                FrmCodigo.Eventos.Nodes.Add NomeMenu5, tvwChild, UCase(FrmTela(TelaAtual).Tag + "." + M.Caption + Trim(Str(M.Ident))), M.Caption, 5
            ElseIf M.Ident = 6 Then
                FrmCodigo.Eventos.Nodes.Add NomeMenu6, tvwChild, UCase(FrmTela(TelaAtual).Tag + "." + M.Caption + Trim(Str(M.Ident))), M.Caption, 5
            End If
            If M.Ident = 0 Then
                NomeMenu1 = UCase(FrmTela(TelaAtual).Tag + "." + M.Caption + Trim(Str(M.Ident)))
            ElseIf M.Ident = 1 Then
                NomeMenu2 = UCase(FrmTela(TelaAtual).Tag + "." + M.Name + Trim(Str(M.Ident)))
            ElseIf M.Ident = 2 Then
                NomeMenu3 = UCase(FrmTela(TelaAtual).Tag + "." + M.Name + Trim(Str(M.Ident)))
            ElseIf M.Ident = 3 Then
                NomeMenu4 = UCase(FrmTela(TelaAtual).Tag + "." + M.Name + Trim(Str(M.Ident)))
            ElseIf M.Ident = 4 Then
                NomeMenu5 = UCase(FrmTela(TelaAtual).Tag + "." + M.Name + Trim(Str(M.Ident)))
            ElseIf M.Ident = 5 Then
                NomeMenu6 = UCase(FrmTela(TelaAtual).Tag + "." + M.Name + Trim(Str(M.Ident)))
            End If
            FrmTela(X).xMenu.MenuTree.Add M
            Rs.MoveNext
        Loop
        FrmTela(X).xMenu.Refresh
        FrmTela(X).SetFocus
ProximaTela1:
        TabelaIndex = TabelaIndex + 1
        X = X + 1
    Loop
    BuscaEventosArq ""
    Abilita True
End If
FrmPrincipal.Prog.Nodes(1).Expanded = True
If CRun.OpenExe = 0 Then Me.Visible = False
End Sub

Public Sub StopTela()
On Error Resume Next
Abilita True, True
Dim X As Long
For X = 0 To ContTela
    Unload Run(X)
Next X
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
F1.Visible = False
F2.Visible = False
If Button.Index = 1 Then
    F1.Visible = True
ElseIf Button.Index = 2 Then
    F2.Visible = True
End If
End Sub





Private Sub Txt_Change()
If Grid.TextMatrix(RowT, 1) = "Legenda" Then
    NovoObj.Caption = Txt.Text
ElseIf Grid.TextMatrix(RowT, 1) = "Texto" Then
    NovoObj.Text = Txt.Text
End If
End Sub

Private Sub Txt_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Grid.TextMatrix(RowT, 2) = Txt.Text
    Grid.SetFocus
    MontaPropriedade
End If
End Sub

Private Sub Txt_LostFocus()
Txt.Visible = False
End Sub

Private Sub MontaPropriedade()
On Error Resume Next
Dim Texto As String, X As Long, Tipo As Boolean

With NovoObj
    For X = 1 To Grid.Rows - 1
        Texto = Grid.TextMatrix(X, 2)
        Select Case UCase(Grid.TextMatrix(X, 1))
            Case "NOME"
                .Tag = Texto
            Case "BORDA"
                .BorderStyle = IIf(Left(Texto, 1) = "S", 1, 0)
            Case "OBJETO 3D"
                .Appearance = IIf(Left(Texto, 1) = "S", 1, 0)
            Case "COR FUNDO"
                .BackColor = Texto
            Case "COR DA LETRA"
                .ForeColor = Texto
            Case "TEXTO"
                .Text = Texto
            Case "MASCARA"
                .PasswordChar = Texto
            Case "TEXTO"
                .Text = Texto
            Case "IMAGEM"
                .ToolTipText = Texto
            Case "COMPRIMIR"
                .Stretch = IIf(Left(Texto, 1) = "S", True, False)
            Case "LEDENDA"
                .Caption = Texto
            Case "BOTAO MAX"
            
            Case "BOTAO MIN"
            Case "BOTAO FECHAR"
            Case "ORDEM"
                .TabIndex = Texto
            Case "TAMX"
                .Height = Texto
            Case "TAMY"
                .Width = Texto
            Case "POXX"
                .Top = Texto
            Case "POXY"
                .Left = Texto
            Case "FONTE"
                .FontName = Texto
            Case "TAMANHO"
                .FontSize = Texto
    
            Case "TEMPO"
                .DataMember = Texto
        End Select
    Next X
End With
        
        
        
End Sub
Private Sub Prog_DblClick()
Dim a As String, X As Long
If Prog.Nodes.Count <> 0 Then
    a = Prog.SelectedItem
    If Trim(a) <> "" Then
        For X = 0 To ContTela - 1
            If UCase(a) = UCase(FrmTela(X).Tag) Then
                Set NovoObj = FrmTela(X)
                TelaAtual = X
                FrmTela(X).Visible = True
                FrmTela(X).SetFocus
                Exit Sub
            End If
        Next X
    End If
End If
End Sub
