VERSION 5.00
Begin VB.Form FrmApresentacao 
   BackColor       =   &H80000009&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Lego - Apresentação"
   ClientHeight    =   3348
   ClientLeft      =   48
   ClientTop       =   288
   ClientWidth     =   5928
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3348
   ScaleWidth      =   5928
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   4320
      TabIndex        =   5
      Top             =   2700
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   2880
      TabIndex        =   6
      Top             =   2700
      Width           =   1455
   End
   Begin VB.TextBox TxtTitulo 
      Height          =   285
      Left            =   1080
      TabIndex        =   0
      Top             =   360
      Width           =   4665
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000009&
      Caption         =   "Plataforma"
      ForeColor       =   &H8000000D&
      Height          =   675
      Left            =   1080
      TabIndex        =   8
      Top             =   1920
      Width           =   4695
      Begin VB.OptionButton O 
         Appearance      =   0  'Flat
         BackColor       =   &H80000009&
         Caption         =   "Win 32"
         ForeColor       =   &H8000000D&
         Height          =   225
         Index           =   1
         Left            =   2910
         TabIndex        =   4
         Top             =   300
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton O 
         Appearance      =   0  'Flat
         BackColor       =   &H80000009&
         Caption         =   "Win 16"
         ForeColor       =   &H8000000D&
         Height          =   225
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   300
         Width           =   1215
      End
   End
   Begin VB.TextBox TxtAutor 
      Height          =   285
      Left            =   1050
      TabIndex        =   1
      Top             =   930
      Width           =   4665
   End
   Begin VB.TextBox TxtSenha 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1050
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1560
      Width           =   4665
   End
   Begin VB.CommandButton CmdAbrir 
      Caption         =   "&Abrir Projeto"
      Height          =   375
      Left            =   1440
      TabIndex        =   7
      Top             =   2700
      Width           =   1455
   End
   Begin VB.Image Image2 
      Height          =   384
      Left            =   36
      Picture         =   "FrmApresentacao.frx":0000
      Top             =   2820
      Width           =   384
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Titulo do Sistema"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   1110
      TabIndex        =   11
      Top             =   150
      Width           =   1665
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Autor"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   1110
      TabIndex        =   10
      Top             =   720
      Width           =   1665
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Senha"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   1080
      TabIndex        =   9
      Top             =   1320
      Width           =   1665
   End
   Begin VB.Image Image1 
      Height          =   1920
      Left            =   0
      Picture         =   "FrmApresentacao.frx":08CA
      Stretch         =   -1  'True
      Top             =   180
      Width           =   1170
   End
   Begin VB.Menu Mnu 
      Caption         =   "M"
      Visible         =   0   'False
      Begin VB.Menu Xmenu 
         Caption         =   ""
         Index           =   0
      End
   End
End
Attribute VB_Name = "FrmApresentacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdAbrir_Click()
NoOpen = 1
Unload Me
End Sub

Private Sub Command1_Click()
Dim NomeProg As String
Titulo_Pjt = TxtTitulo.Text
FrmPrincipal.Caption = "Lego 1.1 [" + Titulo_Pjt + "]"
Autor = TxtAutor.Text
Senha = TxtSenha.Text

If O(0).Value = True Then
    Plataforma = O(0).Caption
Else
    Plataforma = O(1).Caption
End If
NoOpen = 2
Unload Me
End Sub

Private Sub Command2_Click()
NoOpen = 0
Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
If KeyCode = 27 Then Command2_Click
End Sub
Private Sub Xmenu_Click(Index As Integer)
If TipoMenu = True Then
    FrmCodigo.TxtCod.SelText = "." + Xmenu(Index).Caption
Else
    FrmCodigo.TxtCod.SelText = Xmenu(Index).Caption
End If
End Sub
