VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmOpcoes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Inicialição do Programa"
   ClientHeight    =   3540
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   6624
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   6624
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog Com 
      Left            =   5220
      Top             =   2370
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3315
      Left            =   30
      TabIndex        =   11
      Top             =   90
      Width           =   3105
      _ExtentX        =   5461
      _ExtentY        =   5842
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      ForeColor       =   -2147483646
      TabCaption(0)   =   "Inicialização"
      TabPicture(0)   =   "FrmOpcoes.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "List"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Icones"
      TabPicture(1)   =   "FrmOpcoes.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label5"
      Tab(1).Control(1)=   "I1"
      Tab(1).Control(2)=   "Label6"
      Tab(1).Control(3)=   "I2"
      Tab(1).Control(4)=   "Label7"
      Tab(1).Control(5)=   "Image1"
      Tab(1).Control(6)=   "CboIcones"
      Tab(1).Control(7)=   "Command3"
      Tab(1).Control(8)=   "Command4"
      Tab(1).ControlCount=   9
      TabCaption(2)   =   "Opções"
      TabPicture(2)   =   "FrmOpcoes.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "ChkFechar"
      Tab(2).Control(1)=   "ChkEnter"
      Tab(2).ControlCount=   2
      Begin VB.CheckBox ChkEnter 
         Caption         =   "Pular automaticamente para proximo objeto  quando a Tecla ""Enter"" for apertada"
         ForeColor       =   &H8000000D&
         Height          =   795
         Left            =   -74820
         TabIndex        =   21
         Top             =   990
         Width           =   2835
      End
      Begin VB.CheckBox ChkFechar 
         Caption         =   "Fechar tela ao Tecla ""ESC"""
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   -74820
         TabIndex        =   20
         Top             =   600
         Width           =   2655
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Trocar"
         Height          =   435
         Left            =   -73080
         TabIndex        =   19
         Top             =   2730
         Width           =   975
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Procurar"
         Height          =   435
         Left            =   -74040
         TabIndex        =   18
         Top             =   2730
         Width           =   975
      End
      Begin VB.ComboBox CboIcones 
         Height          =   288
         Left            =   -74940
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   810
         Width           =   2895
      End
      Begin VB.ListBox List 
         BackColor       =   &H0080C0FF&
         Height          =   2544
         Left            =   90
         TabIndex        =   12
         Top             =   600
         Width           =   2895
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   705
         Left            =   -74790
         Picture         =   "FrmOpcoes.frx":0054
         Stretch         =   -1  'True
         Top             =   2310
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Novo"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   -74010
         TabIndex        =   17
         Top             =   1230
         Width           =   405
      End
      Begin VB.Image I2 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   705
         Left            =   -74040
         Picture         =   "FrmOpcoes.frx":091E
         Stretch         =   -1  'True
         Top             =   1500
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Atual"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   -74880
         TabIndex        =   16
         Top             =   1260
         Width           =   405
      End
      Begin VB.Image I1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   705
         Left            =   -74910
         Stretch         =   -1  'True
         Top             =   1500
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Icones da Tela"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   -74910
         TabIndex        =   15
         Top             =   540
         Width           =   1665
      End
      Begin VB.Label Label1 
         Caption         =   "Iniciar o Programa em :"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   90
         TabIndex        =   13
         Top             =   360
         Width           =   1665
      End
   End
   Begin VB.TextBox TxtSenha 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   3210
      PasswordChar    =   "*"
      TabIndex        =   10
      Top             =   2520
      Width           =   3255
   End
   Begin VB.TextBox TxtAutor 
      Height          =   285
      Left            =   3210
      TabIndex        =   8
      Top             =   1890
      Width           =   3255
   End
   Begin VB.Frame Frame1 
      Caption         =   "Plataforma"
      ForeColor       =   &H8000000D&
      Height          =   885
      Left            =   3210
      TabIndex        =   4
      Top             =   690
      Width           =   3225
      Begin VB.OptionButton O 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Caption         =   "Win 32"
         ForeColor       =   &H8000000D&
         Height          =   225
         Index           =   1
         Left            =   1830
         TabIndex        =   6
         Top             =   360
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton O 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Caption         =   "Win 16"
         ForeColor       =   &H8000000D&
         Height          =   225
         Index           =   0
         Left            =   210
         TabIndex        =   5
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.TextBox TxtTitulo 
      Height          =   285
      Left            =   3210
      TabIndex        =   3
      Top             =   300
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   4800
      TabIndex        =   0
      Top             =   3060
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Top             =   3060
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Senha"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   3210
      TabIndex        =   9
      Top             =   2280
      Width           =   1665
   End
   Begin VB.Label Label3 
      Caption         =   "Autor"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   3210
      TabIndex        =   7
      Top             =   1650
      Width           =   1665
   End
   Begin VB.Label Label2 
      Caption         =   "Titulo do Sistema"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   3210
      TabIndex        =   2
      Top             =   60
      Width           =   1665
   End
End
Attribute VB_Name = "FrmOpcoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CboIcones_Click()
If CboIcones.ListIndex <> -1 Then
    I1.Picture = FrmTela(CboIcones.ListIndex).Icon
End If
End Sub

Private Sub Command1_Click()
Nome_Da_Tela = List.Text
Titulo_Pjt = TxtTitulo.Text
FrmPrincipal.Caption = "Lego 1.0 [" + Titulo_Pjt + "]"
Autor = TxtAutor.Text
Senha = TxtSenha.Text
If O(0).Value = True Then
    Plataforma = O(0).Caption
Else
    Plataforma = O(1).Caption
End If
TeclaEnter = ChkEnter.Value
TeclaEsc = ChkFechar.Value
'LocalBancodeDados = TxtBanco.Text
If Trim(LocalBancodeDados) = "" Then
    FrmPrincipal.T.Buttons(10).Enabled = False
Else
    FrmPrincipal.T.Buttons(10).Enabled = True
End If
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
Com.FileName = ""
Com.Filter = "Icones (*.Ico) |*.Ico|Todos os Arquivo (*.*)|*.*"
Com.ShowOpen

If Com.FileName <> "" Then
    I2.Picture = LoadPicture(Com.FileName)
End If
End Sub

Private Sub Command4_Click()
If CboIcones.ListIndex <> -1 Then
    FrmTela(CboIcones.ListIndex).Icon = I2.Picture
    I2.Picture = Image1.Picture
End If

End Sub

Private Sub Command5_Click()
On Error Resume Next
Com.FileName = ""
Com.Filter = "Arquivo do MicroSoft Access 97/2000|*.mdb|Todos os Arquivo |*.*"
Com.FileTitle = "Carregar banco de dados"
Com.ShowOpen
If Trim(Com.FileName) <> "" Then
    TxtBanco.Text = Com.FileName
End If
End Sub

Private Sub Form_Load()
Dim X As Long
List.Clear
For X = 0 To ContTela - 1
    List.AddItem FrmTela(X).Tag
    CboIcones.AddItem FrmTela(X).Tag
Next X
CboIcones.ListIndex = 0
List.Text = Nome_Da_Tela
TxtTitulo.Text = Titulo_Pjt
TxtAutor.Text = Autor
TxtSenha.Text = Senha
If UCase(Plataforma) = "WIN 32" Then
    O(1).Value = True
    O(0).Value = False
Else
    O(1).Value = False
    O(0).Value = True
End If
ChkEnter.Value = TeclaEnter
ChkFechar.Value = TeclaEsc
'TxtBanco.Text = LocalBancodeDados
End Sub

