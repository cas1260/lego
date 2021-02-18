VERSION 5.00
Object = "{83501F6F-CBF0-4D8D-8EA4-9E57E403D680}#1.0#0"; "TOOLBAR3.OCX"
Begin VB.Form FrmR 
   ClientHeight    =   4392
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   7356
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   4392
   ScaleWidth      =   7356
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   840
      Top             =   3150
   End
   Begin VB.ComboBox Cbo 
      Height          =   315
      Index           =   0
      Left            =   2610
      TabIndex        =   6
      Top             =   720
      Visible         =   0   'False
      Width           =   2445
   End
   Begin VB.Frame Fm 
      Height          =   525
      Index           =   0
      Left            =   4710
      TabIndex        =   4
      Top             =   3000
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.TextBox Txt 
      Height          =   285
      Index           =   0
      Left            =   5040
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   3810
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.CommandButton Cmd 
      Height          =   525
      Index           =   0
      Left            =   2070
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2700
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.CheckBox Chk 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   5190
      TabIndex        =   1
      Top             =   1710
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.ListBox Lst 
      Height          =   432
      Index           =   0
      ItemData        =   "FrmR1.frx":0000
      Left            =   1980
      List            =   "FrmR1.frx":0002
      TabIndex        =   0
      Top             =   1770
      Visible         =   0   'False
      Width           =   1965
   End
   Begin ctlToolBar.xMenu xMenu 
      Align           =   1  'Align Top
      Height          =   345
      Left            =   0
      Negotiate       =   -1  'True
      TabIndex        =   7
      Top             =   0
      Width           =   7350
      _ExtentX        =   12975
      _ExtentY        =   614
      BeginProperty ItemsFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Botao 
      Caption         =   "111"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   1080
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Lbl 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H80000008&
      Height          =   165
      Index           =   0
      Left            =   4110
      TabIndex        =   5
      Top             =   2430
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.Image Img 
      BorderStyle     =   1  'Fixed Single
      Height          =   555
      Index           =   0
      Left            =   3000
      Top             =   3720
      Visible         =   0   'False
      Width           =   675
   End
End
Attribute VB_Name = "FrmR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cbo_Click(Index As Integer)
Rodar BuscaEventos(Me.Tag, Cbo(Index).Tag, "1")
End Sub

Private Sub Cbo_DblClick(Index As Integer)
Rodar BuscaEventos(Me.Tag, Cbo(Index).Tag, "2")
End Sub

Private Sub Cbo_GotFocus(Index As Integer)
Rodar BuscaEventos(Me.Tag, Cbo(Index).Tag, "Ganhar")
End Sub

Private Sub Cbo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
TeclaKey = KeyCode
Rodar BuscaEventos(Me.Tag, Cbo(Index).Tag, "Escrever")
End Sub

Private Sub Chk_Click(Index As Integer)
Rodar BuscaEventos(Me.Tag, Chk(Index).Tag, "1")
End Sub

Private Sub Chk_GotFocus(Index As Integer)
Rodar BuscaEventos(Me.Tag, Chk(Index).Tag, "Ganhar")
End Sub

Private Sub Chk_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
TeclaKey = KeyCode
Rodar BuscaEventos(Me.Tag, Chk(Index).Tag, "Escrever")
End Sub

Private Sub Cmd_Click(Index As Integer)
Rodar BuscaEventos(Me.Tag, Cmd(Index).Tag, "1")
End Sub

Private Sub Cmd_GotFocus(Index As Integer)
Rodar BuscaEventos(Me.Tag, Cmd(Index).Tag, "Ganhar")
End Sub

Private Sub Cmd_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
TeclaKey = KeyCode
Rodar BuscaEventos(Me.Tag, Cmd(Index).Tag, "Escrever")
End Sub

Private Sub Cmd_LostFocus(Index As Integer)
Rodar BuscaEventos(Me.Tag, Cmd(Index).Tag, "Perder")
End Sub

Private Sub Fm_Click(Index As Integer)
Rodar BuscaEventos(Me.Tag, Fm(Index).Tag, "1")
End Sub

Private Sub Fm_DblClick(Index As Integer)
Rodar BuscaEventos(Me.Tag, Fm(Index).Tag, "2")
End Sub

Private Sub Form_Activate()
Set FrmTelaRun = Me
If xMenu.MenuTree.Count = 0 Then
'    Form_Load
End If
If Left(Botao.Caption, 1) = "0" Then
    ToggleSysMenuEnableDisable Me.HWnd, SMSC_CLOSE
End If
If Mid(Botao.Caption, 2, 1) = "0" Then
   ToggleSysMenuEnableDisable Me.HWnd, SMSC_MAXIMIZE
End If
If Right(Botao.Caption, 1) = "0" Then
   ToggleSysMenuEnableDisable Me.HWnd, SMSC_MINIMIZE
End If
End Sub

Private Sub Form_Click()
Set FrmTelaRun = Me
Rodar BuscaEventos(Me.Tag, "1")
End Sub

Private Sub Form_DblClick()
Rodar BuscaEventos(Me.Tag, "2")
End Sub

Private Sub Form_GotFocus()
Set FrmTelaRun = Me
Rodar BuscaEventos(Me.Tag, "Ganhar")
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If TeclaEnter = 1 Then
        SendKeys "{TAB}"
    End If
End If
If KeyCode = 27 Then
    If TeclaEsc = 1 Then
        Unload Me
    End If
End If
'Rodar BuscaEventos(Me.Tag, "Escrever")
End Sub

Private Sub Form_LostFocus()
Rodar BuscaEventos(Me.Tag, "Perder")
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Rodar BuscaEventos(Me.Tag, "Fechar")
End Sub
Private Sub Form_Resize()
Rodar BuscaEventos(Me.Tag, "Red", "")
End Sub

Private Sub Img_Click(Index As Integer)
Rodar BuscaEventos(Me.Tag, Img(Index).Tag, "1")
End Sub

Private Sub Img_DblClick(Index As Integer)
Rodar BuscaEventos(Me.Tag, Img(Index).Tag, "2")
End Sub

Private Sub Lbl_Click(Index As Integer)
Rodar BuscaEventos(Me.Tag, Lbl(Index).Tag, "1")
End Sub

Private Sub Lbl_DblClick(Index As Integer)
Rodar BuscaEventos(Me.Tag, Lbl(Index).Tag, "2")
End Sub

Private Sub Lst_Click(Index As Integer)
Rodar BuscaEventos(Me.Tag, Lst(Index).Tag, "1")
End Sub

Private Sub Lst_DblClick(Index As Integer)
Rodar BuscaEventos(Me.Tag, Lst(Index).Tag, "2")
End Sub

Private Sub Lst_GotFocus(Index As Integer)
Rodar BuscaEventos(Me.Tag, Lst(Index).Tag, "Ganhar")
End Sub

Private Sub Lst_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
TeclaKey = KeyCode
Rodar BuscaEventos(Me.Tag, Lst(Index).Tag, "Escrever")
End Sub

Private Sub Timer1_Timer()
'Set FrmTelaRun = Nothing
'Set FrmTelaRun = Me
Timer1.Interval = 0
Timer1.Enabled = False

End Sub

Private Sub Txt_Click(Index As Integer)
Rodar BuscaEventos(Me.Tag, Txt(Index).Tag, "1")
End Sub

Private Sub Txt_DblClick(Index As Integer)
Rodar BuscaEventos(Me.Tag, Txt(Index).Tag, "1")
End Sub

Private Sub Txt_GotFocus(Index As Integer)
Txt(Index).SelStart = 0
Txt(Index).SelLength = Len(Txt(Index).Text)
Rodar BuscaEventos(Me.Tag, Txt(Index).Tag, "Ganhar")
End Sub

Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
TeclaKey = KeyCode
Rodar BuscaEventos(Me.Tag, Txt(Index).Tag, "Escrever")
End Sub

Private Sub xMenu_ItemClick(Key As String)
X = xMenu.KeyToIndex(Key)
X = xMenu.MenuTree.Item(X).Ident
Rodar BuscaEventos(Me.Tag, Key + Trim(Str(X)), "")
End Sub

