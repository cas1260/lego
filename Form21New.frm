VERSION 5.00
Object = "{83501F6F-CBF0-4D8D-8EA4-9E57E403D680}#1.0#0"; "TOOLBAR3.OCX"
Begin VB.Form Form2 
   BackColor       =   &H00737373&
   ClientHeight    =   5475
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6690
   Icon            =   "Form21New.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   NegotiateMenus  =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   6690
   Visible         =   0   'False
   Begin VB.Data Banco 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Index           =   0
      Left            =   2490
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4980
      Visible         =   0   'False
      Width           =   2475
   End
   Begin ctlToolBar.xMenu xMenu 
      Align           =   1  'Align Top
      Height          =   345
      Left            =   0
      Negotiate       =   -1  'True
      TabIndex        =   18
      Top             =   0
      Width           =   6690
      _ExtentX        =   11800
      _ExtentY        =   609
      Style           =   1
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
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   960
      Top             =   3600
   End
   Begin VB.PictureBox Im 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   90
      Index           =   7
      Left            =   480
      ScaleHeight     =   60
      ScaleWidth      =   60
      TabIndex        =   17
      Top             =   2460
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.PictureBox Im 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   90
      Index           =   6
      Left            =   510
      ScaleHeight     =   60
      ScaleWidth      =   60
      TabIndex        =   16
      Top             =   2820
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.PictureBox Im 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   90
      Index           =   5
      Left            =   1440
      ScaleHeight     =   60
      ScaleWidth      =   60
      TabIndex        =   15
      Top             =   2460
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.PictureBox Im 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   90
      Index           =   4
      Left            =   930
      ScaleHeight     =   60
      ScaleWidth      =   60
      TabIndex        =   14
      Top             =   2580
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.PictureBox Im 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   90
      Index           =   3
      Left            =   1440
      ScaleHeight     =   60
      ScaleWidth      =   60
      TabIndex        =   13
      Top             =   2940
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.PictureBox Im 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   90
      Index           =   2
      Left            =   990
      ScaleHeight     =   60
      ScaleWidth      =   60
      TabIndex        =   12
      Top             =   3000
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.PictureBox Im 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   90
      Index           =   1
      Left            =   930
      ScaleHeight     =   60
      ScaleWidth      =   60
      TabIndex        =   11
      Top             =   2280
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.PictureBox Im 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   90
      Index           =   0
      Left            =   300
      ScaleHeight     =   60
      ScaleWidth      =   60
      TabIndex        =   10
      Top             =   2160
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.ListBox Lst 
      BackColor       =   &H00FFFFFF&
      Height          =   840
      Index           =   0
      ItemData        =   "Form21New.frx":08CA
      Left            =   750
      List            =   "Form21New.frx":08CC
      TabIndex        =   9
      Top             =   1590
      Visible         =   0   'False
      Width           =   1965
   End
   Begin VB.TextBox TxtM 
      Height          =   1035
      Index           =   0
      Left            =   1500
      MousePointer    =   1  'Arrow
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   3480
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.TextBox Focus 
      Height          =   285
      Left            =   5790
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CheckBox Chk 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   3630
      TabIndex        =   4
      Top             =   1620
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.CommandButton Cmd 
      BackColor       =   &H00737373&
      Height          =   525
      Index           =   0
      Left            =   2220
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2790
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.TextBox Txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Index           =   0
      Left            =   3810
      MousePointer    =   1  'Arrow
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   3660
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.Frame Fm 
      BackColor       =   &H00737373&
      Height          =   525
      Index           =   0
      Left            =   3510
      TabIndex        =   2
      Top             =   2850
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.Image BancoImg 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   660
      Index           =   0
      Left            =   4980
      Picture         =   "Form21New.frx":08CE
      Stretch         =   -1  'True
      Top             =   2100
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Label O 
      Height          =   405
      Left            =   270
      TabIndex        =   19
      Top             =   630
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.Image Img 
      BorderStyle     =   1  'Fixed Single
      Height          =   555
      Index           =   0
      Left            =   4380
      Top             =   900
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Shape S 
      BorderStyle     =   3  'Dot
      Height          =   165
      Left            =   1530
      Top             =   1320
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Image Cbo 
      Height          =   315
      Index           =   0
      Left            =   1200
      Picture         =   "Form21New.frx":1C18
      Stretch         =   -1  'True
      Top             =   1110
      Visible         =   0   'False
      Width           =   2100
   End
   Begin VB.Label Nome 
      Height          =   345
      Left            =   150
      TabIndex        =   6
      Top             =   4080
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Label Cont 
      Height          =   165
      Left            =   3600
      TabIndex        =   5
      Top             =   4560
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.Label Lbl 
      Appearance      =   0  'Flat
      BackColor       =   &H00737373&
      ForeColor       =   &H80000008&
      Height          =   165
      Index           =   0
      Left            =   2910
      TabIndex        =   0
      Top             =   2310
      Visible         =   0   'False
      Width           =   705
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const WM_NCLBUTTONDOWN As Long = &HA1&
Private Const HTCAPTION As Long = 2&

Dim CmdX, CmdY As Long, CmdTipo As Boolean
Dim TxtX, TxtY As Long, TxtTipo As Boolean
Dim FmX, FmY As Long, FmTipo As Boolean
Dim LblX, LblY As Long, LblTipo As Boolean
Dim CboX, CboY As Long, CboTipo As Boolean
'Dim CmdX, CmdY As Long, CmdTipo As Boolean
Dim MenuX, MenuY As Long, MenuTipo As Boolean
Dim Tamanho As Integer
Dim FormX, FormY As Long, FormTipo As Boolean
Dim AcessoObj As String
Dim BordaX() As Long, BordaY() As Long, TipoBorda() As Boolean
Dim TabAux As String
Dim ChkY As Long, ChkX As Long, ChkTipo As Boolean
Dim TotalObj As Totais
Dim ImgX As Integer, ImgY As Integer, ImgTipo As Boolean
Dim LstX As Integer, LstY As Integer, LstTipo As Boolean
Dim BancoX As Integer, BancoY As Integer, BancoTipo As Boolean
Dim Click2Var As Long

Private Sub BancoImg_Click(Index As Integer)
Set NovoObj = BancoImg(Index)
Redimesiona
Selecione1
End Sub

Private Sub Cbo_DblClick(Index As Integer)
Click2 Cbo(Index).Tag
End Sub

Private Sub Cbo_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
CboTipo = True
CboX = X
CboY = Y
End Sub
Private Sub cbo_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If CboTipo And Me.MousePointer = 0 Then
    Cbo(Index).Move Cbo(Index).Left + X - CboX, Cbo(Index).Top + Y - CboY
    Redimesiona
End If
End Sub
Private Sub Cbo_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
CboTipo = False
End Sub
Private Sub Chk_GotFocus(Index As Integer)
Click2Var = 0
End Sub

Private Sub Chk_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Me.MousePointer = 1 Then
    If Button = 1 Then
        Chk_Click Index
        Set NovoObj = Chk(Index)
        Call ReleaseCapture
        Call SendMessage(NovoObj.hwnd, WM_NCLBUTTONDOWN, ByVal HTCAPTION, ByVal 0&)
        m_bMoving = True
        Redimesiona
    End If
End If
End Sub

Private Sub Fm_DblClick(Index As Integer)
Click2 Fm(Index).Tag
End Sub

Private Sub Form_DblClick()
On Error Resume Next
Click2 ""
End Sub

Private Sub Form_GotFocus()
Form_Click
TelaAtual = IIf(Trim(Cont.Caption) = "", 0, Cont.Caption)
End Sub

Private Sub Img_DblClick(Index As Integer)
Click2 Img(Index).Tag
End Sub

Private Sub Img_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

End Sub

Private Sub Lbl_DblClick(Index As Integer)
Click2 Lbl(Index).Tag
End Sub

Private Sub Lst_Click(Index As Integer)
On Error Resume Next
FrmPropriedade.CboObj.Text = Lst(Index).Tag
End Sub

Private Sub Lst_DblClick(Index As Integer)
Click2 Lst(Index).Tag
End Sub

Private Sub Lst_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If LstTipo And Me.MousePointer = 0 Then
    Lst(Index).Move Lst(Index).Left + X - LstX, Lst(Index).Top + Y - LstY
    Redimesiona
ElseIf Me.MousePointer = 2 And Button = 1 Then
    Form_MouseMove Button, Shift, X, Y
End If
End Sub

Private Sub Lst_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
LstTipo = False
If Me.MousePointer = 1 And Button = 1 Then
    Form_MouseDown Button, Shift, X, Y
End If
If NovoObj.Name <> "Lst" Then
    Set NovoObj = Lst(Index)
End If
Selecione1
Redimesiona
End Sub
Private Sub Lst_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 And Me.MousePointer = 1 Then
    Set NovoObj = Lst(Index)
    Selecione1
    Redimesiona
'ElseIf Me.MousePointer = 2 And Button = 1 Then
    'Form_MouseDown Button, Shift, X, Y
End If
LstY = Y
LstX = X
LstTipo = True
End Sub

Private Sub Chk_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If ChkTipo Then
    Chk(Index).Move Chk(Index).Left + X - ChkX, Chk(Index).Top + Y - ChkY
    Click2Var = 0
    Redimesiona
End If
End Sub

Private Sub Chk_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
ChkTipo = False
End Sub

Private Sub Fm_Click(Index As Integer)
FrmPropriedade.CboObj.Text = Fm(Index).Tag
Set NovoObj = Fm(Index)
Redimesiona
Selecione1
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = vbKeyDelete Then
    If UCase(NovoObj.Name) = "FORM2" Then
        Exit Sub
    End If
    Dim Xp As Long
    For Xp = 1 To FrmCodigo.Eventos.Nodes.Count
        If UCase(FrmCodigo.Eventos.Nodes(Xp).Text) = UCase(NovoObj.Tag) Then
            FrmCodigo.Eventos.Nodes.Remove Xp
            Exit For
        End If
    Next Xp
    Unload NovoObj
    If Err = 340 Then
        Exit Sub
    End If
    ' MsgBox
    Im(0).Visible = False
    Im(1).Visible = False
    Im(2).Visible = False
    Im(3).Visible = False
    Im(4).Visible = False
    Im(5).Visible = False
    Im(6).Visible = False
    Im(7).Visible = False
Else
    FrmPropriedade.Grid.SetFocus
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Me.MousePointer = 2 Then
    If FormTipo Then
        'Me.Caption = x & "   " & Y
        If Y > S.Top Then
            S.Height = Y - FormY
        End If
        If X > S.Left Then
            S.Width = X - FormX
        End If
    End If
End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormTipo = False
S.Visible = False

If FrmPrincipal.T.Buttons(2).Value = tbrPressed Then
    
    CrieObj 0
    Me.MousePointer = 1
    FrmPrincipal.T.Buttons(2).Value = tbrUnpressed

ElseIf FrmPrincipal.T.Buttons(3).Value = tbrPressed Then
    
    CrieObj 1
    Me.MousePointer = 1
    FrmPrincipal.T.Buttons(3).Value = tbrUnpressed

ElseIf FrmPrincipal.T.Buttons(4).Value = tbrPressed Then
    
    CrieObj 2
    Me.MousePointer = 1
    FrmPrincipal.T.Buttons(4).Value = tbrUnpressed
    
ElseIf FrmPrincipal.T.Buttons(5).Value = tbrPressed Then
    
    CrieObj 3
    Me.MousePointer = 1
    FrmPrincipal.T.Buttons(5).Value = tbrUnpressed
    
ElseIf FrmPrincipal.T.Buttons(6).Value = tbrPressed Then
    CrieObj 4
    Me.MousePointer = 1
    FrmPrincipal.T.Buttons(6).Value = tbrUnpressed
    
ElseIf FrmPrincipal.T.Buttons(7).Value = tbrPressed Then
    CrieObj 6
    Me.MousePointer = 1
    FrmPrincipal.T.Buttons(7).Value = tbrUnpressed

ElseIf FrmPrincipal.T.Buttons(8).Value = tbrPressed Then
    CrieObj 5
    Me.MousePointer = 1
    FrmPrincipal.T.Buttons(8).Value = tbrUnpressed
ElseIf FrmPrincipal.T.Buttons(9).Value = tbrPressed Then
    CrieObj 7
    Me.MousePointer = 1
    FrmPrincipal.T.Buttons(9).Value = tbrUnpressed
ElseIf FrmPrincipal.T.Buttons(10).Value = tbrPressed Then
    CrieObj 8
    Me.MousePointer = 1
    FrmPrincipal.T.Buttons(10).Value = tbrUnpressed
End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'    ReadyToClose = True

 '   Cancel = Not ReadyToClose
End Sub

Private Sub Cbo_Click(Index As Integer)
FrmPropriedade.CboObj.Text = Cbo(Index).Tag
Set NovoObj = Cbo(Index)
Selecione1
Redimesiona
End Sub


Private Sub Chk_Click(Index As Integer)
FrmPropriedade.CboObj.Text = Chk(Index).Tag
Chk(Index).Value = 0
Set NovoObj = Chk(Index)
IndexObj = Index
Redimesiona
Selecione1
Click2Var = Click2Var + 1
If Click2Var = 2 Then
    'Click2 Chk(Index).Tag
End If
End Sub

Private Sub Cmd_Click(Index As Integer)
On Error Resume Next
Click2Var = Click2Var + 1
If Click2Var = 2 Then
    Click2 Cmd(Index).Tag
End If
FrmPropriedade.CboObj.Text = Cmd(Index).Tag
Set NovoObj = Cmd(Index)
Redimesiona
Selecione1
End Sub

Private Sub Form_Activate()
On Error Resume Next
If Cont.Caption <> "" Then
    Selecione1
'    Menu.Chama Menus(Cont.Caption)
'    Me.SetFocus
End If
Me.ScaleMode = 1
End Sub

Private Sub Im_Click(Index As Integer)
'MoveAdd Index
End Sub

Private Sub Im_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
TipoBorda(Index) = True
BordaX(Index) = X
BordaY(Index) = Y
End Sub

Private Sub Im_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    On Error Resume Next
    If TipoBorda(Index) Then
        If Index = 0 Then
            Im(0).Move Im(0).Left + X - BordaX(Index), Im(0).Top + Y - BordaY(Index)
            NovoObj.Height = NovoObj.Height - (Im(0).Top - NovoObj.Top)
            NovoObj.Width = NovoObj.Width - (Im(0).Left - NovoObj.Left)
            NovoObj.Top = Im(0).Top
            NovoObj.Left = Im(0).Left
        ElseIf Index = 1 Then
            Im(0).Move Im(0).Left, Im(0).Top + Y - BordaY(Index)
            NovoObj.Height = NovoObj.Height - (Im(0).Top - NovoObj.Top)
        '    NovoObj.Width = NovoObj.Width - (Im(0).Left - NovoObj.Left)
            NovoObj.Top = Im(0).Top
         '   NovoObj.Left = Im(0).Left
        ElseIf Index = 2 Then
            Im(2).Move Im(2).Left + X - BordaX(Index), Im(2).Top + Y - BordaY(Index)
            NovoObj.Height = NovoObj.Height - (Im(2).Top - NovoObj.Top)
            NovoObj.Top = Im(2).Top
            NovoObj.Width = Im(2).Left - Im(0).Left
            'NovoObj.Left = Im(2).Left - NovoObj.Width
        ElseIf Index = 3 Then
            Im(3).Move Im(3).Left + X - BordaX(Index), Im(3).Top + Y - BordaY(Index)
            NovoObj.Width = NovoObj.Width + (NovoObj.Left - Im(3).Left)    'Im(3).Left + NovoObj.Left
            NovoObj.Left = Im(3).Left
            NovoObj.Height = Im(3).Top - NovoObj.Top
        ElseIf Index = 4 Then
            Im(4).Move Im(4).Left + X - BordaX(Index), Im(4).Top + Y - BordaY(Index)
            NovoObj.Height = Im(4).Top - NovoObj.Top
        ElseIf Index = 5 Then
            Im(5).Move Im(5).Left + X - BordaX(Index), Im(5).Top + Y - BordaY(Index)
            If (Im(5).Top - NovoObj.Top) > 0 Then
                NovoObj.Height = Im(5).Top - NovoObj.Top
            End If
            If (Im(5).Left - NovoObj.Left) > 0 Then
                NovoObj.Width = (Im(5).Left - NovoObj.Left)
            End If
        ElseIf Index = 6 Then
            Im(6).Move Im(6).Left + X - BordaX(Index), Im(6).Top + Y - BordaY(Index)
            NovoObj.Width = NovoObj.Width - (Im(6).Left - NovoObj.Left)
            NovoObj.Left = Im(6).Left
        ElseIf Index = 7 Then
            Im(7).Move Im(7).Left + X - BordaX(Index), Im(7).Top + Y - BordaY(Index)
            NovoObj.Width = Im(7).Left - NovoObj.Left
        End If
        Redimesiona
        Selecione1
    End If
End If
End Sub

Private Sub Im_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'TipoBorda(Index) = False
'If NovoObj.Name = "Form2" Then
'    Selecione 1, CCur(IndexObj)
'ElseIf NovoObj.Name = "Cmd" Then
'    Selecione 0, CCur(IndexObj)
'End If
Selecione1
End Sub

Private Sub Img_Click(Index As Integer)
On Error Resume Next
FrmPropriedade.CboObj.Text = Img(Index).Tag
Set NovoObj = Img(Index)
Selecione1
Redimesiona
End Sub

Private Sub Img_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
ImgTipo = True
ImgX = X
ImgY = Y
End Sub
Private Sub BancoImg_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
BancoTipo = True
BancoX = X
BancoY = Y
End Sub
Private Sub BancoImg_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If BancoTipo And Me.MousePointer = 0 Then
    BancoImg(Index).Move BancoImg(Index).Left + X - BancoX, BancoImg(Index).Top + Y - BancoY
    Redimesiona
End If
End Sub
Private Sub BancoImg_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
BancoTipo = False
Redimesiona
End Sub

Private Sub Cmd_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Me.MousePointer = 1 Then
    If Button = 1 Then
        Cmd_Click Index
        Set NovoObj = Cmd(Index)
        Call ReleaseCapture
        Call SendMessage(Cmd(Index).hwnd, WM_NCLBUTTONDOWN, ByVal HTCAPTION, ByVal 0&)
        m_bMoving = True
        Redimesiona
    End If
End If
End Sub

Private Sub Form_Resize()
If Me.WindowState = 1 Then
    Exit Sub
End If
If Cont.Caption <> "" Then
    TelaAtual = Cont.Caption
End If
If Cont.Caption <> "" Then
    'Menus(Cont.Caption).Width = Me.ScaleWidth
End If
Focus.Top = Me.ScaleHeight
Focus.Left = Me.ScaleWidth
If TelaAtual <> -1 Then
    'If FrmPropriedade.Grid.Rows <> 7 Then
    '    Selecione 1, 0
    'End If
    FrmPropriedade.Grid.Row = 10
    FrmPropriedade.Grid.Col = 2
    FrmPropriedade.Grid.Text = FrmTela(TelaAtual).Height
    FrmPropriedade.Grid.Row = 11
    FrmPropriedade.Grid.Col = 2
    FrmPropriedade.Grid.Text = FrmTela(TelaAtual).Width
    'Selecione 1, 0
End If

End Sub
Private Sub Lbl_Click(Index As Integer)
On Error Resume Next
'Lbl(Index).BorderStyle = 1
FrmPropriedade.CboObj.Text = Lbl(Index).Tag
Set NovoObj = Lbl(Index)
Redimesiona
Selecione1
End Sub


Private Sub O_Change()
O_Click
End Sub

Private Sub O_Click()
If O.Caption = "0" Then
    TotalObj.TotalCmd = TotalObj.TotalCmd + 1
ElseIf O.Caption = "1" Then
    TotalObj.TotalFm = TotalObj.TotalFm + 1
ElseIf O.Caption = "2" Then
    TotalObj.TotalImg = TotalObj.TotalImg + 1
ElseIf O.Caption = "3" Then
    TotalObj.TotalLbl = TotalObj.TotalLbl + 1
ElseIf O.Caption = "4" Then
    TotalObj.TotalChk = TotalObj.TotalChk + 1
ElseIf O.Caption = "5" Then
    TotalObj.TotalCombo = TotalObj.TotalCombo + 1
ElseIf O.Caption = "6" Then
    TotalObj.TotalTxt = TotalObj.TotalTxt + 1
ElseIf O.Caption = "7" Then
    TotalObj.TotalLst = TotalObj.TotalLst + 1
ElseIf O.Caption = "6" Then
    TotalObj.TotalOpt = TotalObj.TotalOpt + 1
ElseIf O.Caption = "7" Then
    TotalObj.TotalBanco = TotalObj.TotalBanco + 1
End If
End Sub

Private Sub Timer1_Timer()
Click2Var = 0
End Sub

Private Sub Txt_Click(Index As Integer)
On Error Resume Next
FrmPropriedade.CboObj.Text = Txt(Index).Tag
Set NovoObj = Txt(Index)
Redimesiona
Selecione1
End Sub

Private Sub Txt_DblClick(Index As Integer)
Click2 Txt(Index).Tag
End Sub

Private Sub Txt_GotFocus(Index As Integer)
Focus.SetFocus
End Sub
Private Sub Txt_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If TxtTipo And Me.MousePointer = 0 Then
    Txt(Index).Move Txt(Index).Left + X - TxtX, Txt(Index).Top + Y - TxtY
    Redimesiona
End If
End Sub

Private Sub Txt_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
TxtTipo = False
End Sub
Private Sub Txt_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Me.MousePointer = 1 Then
    If Button = 1 Then
        Txt_Click Index
        Set NovoObj = Txt(Index)
        Call ReleaseCapture
        Call SendMessage(Txt(Index).hwnd, WM_NCLBUTTONDOWN, ByVal HTCAPTION, ByVal 0&)
        m_bMoving = True
        Redimesiona
    End If
End If
End Sub

Private Sub Form_Click()
On Error Resume Next
Set NovoObj = Me
Click2Var = 0
Im(0).Visible = False
Im(1).Visible = False
Im(2).Visible = False
Im(3).Visible = False
Im(4).Visible = False
Im(5).Visible = False
Im(6).Visible = False
Im(7).Visible = False
S.Visible = False
TelaAtual = IIf(Trim(Cont.Caption) = "", 0, Cont.Caption)
Selecione1
'Selecione 1, 0
Focus.SetFocus
If NovoObj.Name = "Form2" Then
    Nome = FrmPropriedade.CboObj.Text
    FrmPropriedade.CboObj.Clear
    FrmPropriedade.CboObj.AddItem Me.Tag
    For Each Ob In Me
        If Ob.Name = "xMenu" Or Ob.Name = "Im" Or Ob.Name = "Focus" Or Ob.Name = "S" Or Ob.Name = "Nome" Or Ob.Name = "Cont" Then
            GoTo proximo
        ElseIf Ob.Index = 0 Then
        Else
            FrmPropriedade.CboObj.AddItem Ob.Tag
        End If
proximo:
    Next
    FrmPropriedade.CboObj.Text = Me.Tag
End If
End Sub

Private Sub Form_Load()
'Banco.DatabaseName = LocalBancodeDados
Me.Visible = False
Me.Width = 7305
ReDim BordaX(7) As Long, BordaY(7) As Long, TipoBorda(7) As Boolean
Im(0).MousePointer = 8
Im(1).MousePointer = 7
Im(2).MousePointer = 6
Im(3).MousePointer = 6
Im(4).MousePointer = 7
Im(5).MousePointer = 8
Im(6).MousePointer = 9
Im(7).MousePointer = 9
TotalObj.TotalChk = 1
TotalObj.TotalCmd = 1
TotalObj.TotalFm = 1
TotalObj.TotalImg = 1
TotalObj.TotalLbl = 1
TotalObj.TotalOpt = 1
TotalObj.TotalTxt = 1
TotalObj.TotalCombo = 1
TotalObj.TotalLst = 1
TotalObj.TotalBanco = 1
Form_Resize

Me.Left = 0
Me.Top = 0
'Me.Visible = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Cancel = 0 Then
    If FechaGeral = True Then
        Me.Visible = False
        Cancel = 1
    End If
End If
End Sub


Private Sub Fm_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Me.MousePointer = 1 Then
    If Button = 1 Then
        Fm_Click Index
        Set NovoObj = Fm(Index)
        Call ReleaseCapture
        Call SendMessage(NovoObj.hwnd, WM_NCLBUTTONDOWN, ByVal HTCAPTION, ByVal 0&)
        m_bMoving = True
        Redimesiona
    End If
End If
End Sub
Private Sub Lbl_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Me.MousePointer = 1 Then
    If Button = 1 Then
        Lbl_Click Index
        Set NovoObj = Lbl(Index)
        Call ReleaseCapture
        Call SendMessage(NovoObj.hwnd, WM_NCLBUTTONDOWN, ByVal HTCAPTION, ByVal 0&)
        m_bMoving = True
        Redimesiona
    End If
End If
End Sub
Private Sub Mnu_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If MenuTipo Then
    Mnu.Move Mnu.Left + X - MenuX, Mnu.Top + Y - MenuY
End If
End Sub

Private Sub Mnu_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
MenuTipo = False
End Sub
Private Sub Mnu_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MenuTipo = True
MenuX = X
MenuY = Y
End Sub

Private Sub Form_MouseDown(Index As Integer, Shift As Integer, X As Single, Y As Single)
S.Visible = True
S.Height = 0
S.Width = 0
S.Left = X
S.Top = Y
FormTipo = True
FormX = X
FormY = Y
'S.Visible = True
End Sub

Private Sub xMenu_ItemClick(Key As String)
Click2 Key + Trim(Str(xMenu.MenuTree(xMenu.KeyToIndex(Key)).Ident)), True
End Sub

Private Sub CrieObj(Index As Byte)
On Error Resume Next
Dim X As Long
If Index = 0 Then  'Se for um CommandBon
    
    Load Cmd(TotalObj.TotalCmd)
    Set NovoObj = FrmTela(TelaAtual).Cmd(TotalObj.TotalCmd)
    NovoObj.Tag = "Botao" + Trim(Str(TotalObj.TotalCmd))
    TotalObj.TotalCmd = TotalObj.TotalCmd + 1
    For X = 0 To FrmTela(TelaAtual).Cmd.Count - 1
        Err.Number = 0
        If UCase(NovoObj.Tag) = UCase(FrmTela(TelaAtual).Cmd(X).Tag) Then
            If Err.Number = 0 Then
                TotalObj.TotalCmd = TotalObj.TotalCmd + 1
            End If
        End If
    Next X
    NovoObj.Tag = "Botao" + Trim(Str(TotalObj.TotalCmd - 1))
ElseIf Index = 1 Then 'Se for um Frame

    Load Fm(TotalObj.TotalFm)
    Set NovoObj = Fm(TotalObj.TotalFm)
    NovoObj.Tag = "Frame" + Trim(Str(TotalObj.TotalFm - 1))
    TotalObj.TotalFm = TotalObj.TotalFm + 1
    For X = 0 To FrmTela(TelaAtual).Fm.Count
        Err.Number = 0
        If UCase(NovoObj.Tag) = FrmTela(TelaAtual).Fm(X).Tag Then
            If Err.Number = 0 Then
                TotalObj.TotalFm = TotalObj.TotalFm + 1
            End If
        End If
    Next X
    NovoObj.Tag = "Frame" + Trim(Str(TotalObj.TotalFm - 1))
ElseIf Index = 2 Then 'Se for um Imagem
    
    Load Img(TotalObj.TotalImg)
    Set NovoObj = Img(TotalObj.TotalImg)
    NovoObj.Tag = "Imagem" + Trim(Str(TotalObj.TotalImg - 1))
    TotalObj.TotalImg = TotalObj.TotalImg + 1
    For X = 0 To FrmTela(TelaAtual).Img.Count
        Err.Number = 0
        If UCase(NovoObj.Tag) = FrmTela(TelaAtual).Img(X).Tag Then
            If Err.Number = 0 Then
                TotalObj.TotalImg = TotalObj.TotalImg + 1
            End If
        End If
    Next X
    NovoObj.Tag = "Imagem" + Trim(Str(TotalObj.TotalImg - 1))

ElseIf Index = 3 Then 'Se For um Label
    
    Load Lbl(TotalObj.TotalLbl)
    Set NovoObj = Lbl(TotalObj.TotalLbl)
    NovoObj.Tag = "Legenda" + Trim(Str(TotalObj.TotalLbl))
    TotalObj.TotalLbl = TotalObj.TotalLbl + 1
    For X = 0 To FrmTela(TelaAtual).Lbl.Count
        Err.Number = 0
        If UCase(NovoObj.Tag) = FrmTela(TelaAtual).Lbl(X).Tag Then
            If Err.Number = 0 Then
                TotalObj.TotalLbl = TotalObj.TotalLbl + 1
            End If
        End If
    Next X
    NovoObj.Tag = "Legenda" + Trim(Str(TotalObj.TotalLbl - 1))
    NovoObj.Caption = NovoObj.Tag
    
ElseIf Index = 4 Then 'Se for uma Check Box
    
    Load Chk(TotalObj.TotalChk)
    Set NovoObj = Chk(TotalObj.TotalChk)
    NovoObj.Tag = "Check" + Trim(Str(TotalObj.TotalChk))
    TotalObj.TotalChk = TotalObj.TotalChk + 1
    For X = 0 To FrmTela(TelaAtual).Chk.Count
        Err.Number = 0
        If UCase(NovoObj.Tag) = FrmTela(TelaAtual).Chk(X).Tag Then
            If Err.Number = 0 Then
                TotalObj.TotalChk = TotalObj.TotalChk + 1
            End If
        End If
    Next X
    NovoObj.Tag = "Check" + Trim(Str(TotalObj.TotalChk - 1))
ElseIf Index = 5 Then 'Se for um Combo Box
    
    Load Cbo(TotalObj.TotalCombo)
    Set NovoObj = Cbo(TotalObj.TotalCombo)
    NovoObj.Tag = "Combo" + Trim(Str(TotalObj.TotalCombo))
    TotalObj.TotalCombo = TotalObj.TotalCombo + 1
    For X = 0 To FrmTela(TelaAtual).Chk.Count
        Err.Number = 0
        If UCase(NovoObj.Tag) = FrmTela(TelaAtual).Cbo(X).Tag Then
            If Err.Number = 0 Then
                TotalObj.TotalCombo = TotalObj.TotalCombo + 1
            End If
        End If
    Next X
    NovoObj.Tag = "Combo" + Trim(Str(TotalObj.TotalCombo - 1))
ElseIf Index = 6 Then

    Load Txt(TotalObj.TotalTxt)
    Set NovoObj = Txt(TotalObj.TotalTxt)
    NovoObj.Tag = "Texto" + Trim(Str(TotalObj.TotalTxt))
    TotalObj.TotalTxt = TotalObj.TotalTxt + 1
    For X = 0 To FrmTela(TelaAtual).Txt.Count
        Err.Number = 0
        If UCase(NovoObj.Tag) = FrmTela(TelaAtual).Txt(X).Tag Then
            If Err.Number = 0 Then
                TotalObj.TotalTxt = TotalObj.TotalTxt + 1
            End If
        End If
    Next X
    NovoObj.Tag = "Texto" + Trim(Str(TotalObj.TotalTxt - 1))
ElseIf Index = 7 Then

    Load Lst(TotalObj.TotalLst)
    Set NovoObj = Lst(TotalObj.TotalLst)
    NovoObj.Tag = "Lista" + Trim(Str(TotalObj.TotalLst))
    TotalObj.TotalLst = TotalObj.TotalLst + 1
    For X = 0 To FrmTela(TelaAtual).Lst.Count
        Err.Number = 0
        If UCase(NovoObj.Tag) = FrmTela(TelaAtual).Lst(X).Tag Then
            If Err.Number = 0 Then
                TotalObj.TotalLst = TotalObj.TotalLst + 1
            End If
        End If
    Next X
ElseIf Index = 8 Then
    Load BancoImg(TotalObj.TotalBanco)
    Set NovoObj = BancoImg(TotalObj.TotalBanco)
    TotalObj.TotalBanco = TotalObj.TotalBanco + 1
    NovoObj.Width = 600
    NovoObj.Height = 600
Else
    Exit Sub
End If

If Index <> 8 Then
    NovoObj.Caption = NovoObj.Tag
    NovoObj.TabIndex = NovoObj.TabIndex - 19
    NovoObj.Height = S.Height
    NovoObj.Width = S.Width
End If
NovoObj.Left = S.Left
NovoObj.Top = S.Top
NovoObj.Visible = True

FrmCodigo.Eventos.Nodes.Add UCase(Me.Tag), tvwChild, UCase(Me.Tag + "." + NovoObj.Tag), NovoObj.Tag, 3


With FrmCodigo.Eventos

    If Index <> 8 Then
        .Nodes.Add UCase(Me.Tag + "." + NovoObj.Tag), tvwChild, UCase(Me.Tag + "." + NovoObj.Tag + ".1"), "Ao Clicar 1 Vezes", 2
        If Index = 7 Or Index = 6 Or Index = 5 Or Index = 3 Or Index = 2 Or Index = 1 Then
            .Nodes.Add UCase(Me.Tag + "." + NovoObj.Tag), tvwChild, UCase(Me.Tag + "." + NovoObj.Tag + ".2"), "Ao Clicar 2 Vezes", 2
        End If
        .Nodes.Add UCase(Me.Tag + "." + NovoObj.Tag), tvwChild, UCase(Me.Tag + "." + NovoObj.Tag + ".Ganhar"), "Ao Ganhar o focu", 2
        .Nodes.Add UCase(Me.Tag + "." + NovoObj.Tag), tvwChild, UCase(Me.Tag + "." + NovoObj.Tag + ".Perder"), "Ao Perder o focu", 2
        If Index = 6 Or 3 Then
            .Nodes.Add UCase(Me.Tag + "." + NovoObj.Tag), tvwChild, UCase(Me.Tag + "." + NovoObj.Tag + ".Escrever"), "Ao Escrever", 2
        End If
    Else
        .Nodes.Add UCase(Me.Tag + "." + NovoObj.Tag), tvwChild, UCase(Me.Tag + "." + NovoObj.Tag + ".Pular"), "Ao Mudar de Registro", 2
    End If
End With
    
    'Set NovoObj = Cmd(Ut)
    NovoObj.ZOrder vbBringToFront
    'Selecione1
    'Redimesiona
    

End Sub


Public Sub Click2(Obj As String, Optional XMnu As Boolean)
Dim X As Long
Dim Comp As String

If Obj = "" Then
    Comp = Me.Tag + ".1"
Else
    If XMnu = True Then
        Comp = Me.Tag + "." + Obj
    Else
        Comp = Me.Tag + "." + Obj + ".1"
    End If
End If

For X = 1 To FrmCodigo.Eventos.Nodes.Count
    If UCase(Comp) = UCase(FrmCodigo.Eventos.Nodes(X).Key) Then
        FrmCodigo.Visible = True
        FrmCodigo.Eventos.Nodes(X).Selected = True
        FrmCodigo.TxtCod.Text = FrmCodigo.Eventos.Nodes(X).Tag
        FrmCodigo.TxtCod.SelStart = Len(FrmCodigo.TxtCod.Text)
        FrmCodigo.TxtCod.SetFocus
        Exit Sub
    End If
Next X
End Sub

Private Sub XpForm_Click()
Set NovoObj = XpForm
Selecione1
Redimesiona
End Sub
Private Sub Img_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If ImgTipo And Me.MousePointer = 0 Then
    Img(Index).Move Img(Index).Left + X - ImgX, Img(Index).Top + Y - ImgY
    Redimesiona
End If
End Sub
