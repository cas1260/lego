VERSION 5.00
Object = "{83501F6F-CBF0-4D8D-8EA4-9E57E403D680}#1.0#0"; "TOOLBAR3.OCX"
Begin VB.Form FrmR 
   ClientHeight    =   5175
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7005
   LinkTopic       =   "Form1"
   ScaleHeight     =   5175
   ScaleWidth      =   7005
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   840
      Top             =   3150
   End
   Begin ctlToolBar.xMenu xMenu 
      Align           =   1  'Align Top
      Height          =   345
      Left            =   0
      Negotiate       =   -1  'True
      TabIndex        =   7
      Top             =   0
      Width           =   7005
      _ExtentX        =   12356
      _ExtentY        =   609
      BeginProperty ItemsFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2940
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
      Height          =   450
      Index           =   0
      ItemData        =   "FrmR.frx":0000
      Left            =   1980
      List            =   "FrmR.frx":0002
      TabIndex        =   0
      Top             =   1740
      Visible         =   0   'False
      Width           =   1965
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
      Left            =   -90
      Top             =   5130
      Visible         =   0   'False
      Width           =   675
   End
End
Attribute VB_Name = "FrmR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmd_Click(Index As Integer)
Rodar BuscaEventos(Me.Tag, Cmd(Index).Tag, "1")
End Sub

Private Sub Cmd_GotFocus(Index As Integer)
Rodar BuscaEventos(Me.Tag, Cmd(Index).Tag, "Ganhar")
End Sub

Private Sub Cmd_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Rodar BuscaEventos(Me.Tag, Cmd(Index).Tag, "Escrever")
End Sub

Private Sub Cmd_LostFocus(Index As Integer)
Rodar BuscaEventos(Me.Tag, Cmd(Index).Tag, "Perder")
End Sub

Private Sub Form_Activate()
Set FrmTelaRun = Me
If xMenu.MenuTree.Count = 0 Then
'    Form_Load
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
Rodar BuscaEventos(Me.Tag, "Escrever")
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

Private Sub Timer1_Timer()
'Set FrmTelaRun = Nothing
'Set FrmTelaRun = Me
Timer1.Interval = 0
Timer1.Enabled = False

End Sub

Private Sub Txt_GotFocus(Index As Integer)
Txt(Index).SelStart = 0
Txt(Index).SelLength = Len(Txt(Index).Text)
End Sub

Private Sub xMenu_ItemClick(Key As String)
X = xMenu.KeyToIndex(Key)
X = xMenu.MenuTree.Item(X).Ident
Rodar BuscaEventos(Me.Tag, Key + Trim(Str(X)), "")
End Sub

