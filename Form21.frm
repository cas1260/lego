VERSION 5.00
Object = "{83501F6F-CBF0-4D8D-8EA4-9E57E403D680}#1.0#0"; "TOOLBAR3.OCX"
Begin VB.Form Form2 
   BackColor       =   &H80000004&
   ClientHeight    =   5475
   ClientLeft      =   2115
   ClientTop       =   1830
   ClientWidth     =   6690
   Icon            =   "Form21.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   NegotiateMenus  =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   6690
   Visible         =   0   'False
   Begin VB.Timer T 
      Index           =   0
      Left            =   3150
      Top             =   3840
   End
   Begin VB.PictureBox S 
      BackColor       =   &H80000001&
      Height          =   828
      Left            =   5232
      ScaleHeight     =   765
      ScaleWidth      =   750
      TabIndex        =   20
      Top             =   3504
      Visible         =   0   'False
      Width           =   804
   End
   Begin ctlToolBar.xMenu xMenu 
      Align           =   1  'Align Top
      Height          =   348
      Left            =   0
      Negotiate       =   -1  'True
      TabIndex        =   18
      Top             =   0
      Width           =   6696
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
      Height          =   645
      Index           =   0
      ItemData        =   "Form21.frx":08CA
      Left            =   768
      List            =   "Form21.frx":08CC
      TabIndex        =   9
      Top             =   1608
      Visible         =   0   'False
      Width           =   1992
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
      Left            =   4848
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   2712
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
      BackColor       =   &H80000004&
      Height          =   480
      Index           =   0
      Left            =   2208
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2808
      Visible         =   0   'False
      Width           =   1248
   End
   Begin VB.TextBox Txt 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   0
      Left            =   1470
      MousePointer    =   1  'Arrow
      TabIndex        =   1
      Top             =   2430
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.Frame Fm 
      BackColor       =   &H80000004&
      Height          =   525
      Index           =   0
      Left            =   3510
      TabIndex        =   2
      Top             =   2850
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.Image Tm 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   ""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   0
      EndProperty
      DataMember      =   "0"
      Height          =   480
      Index           =   0
      Left            =   5580
      Picture         =   "Form21.frx":08CE
      Top             =   1260
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgRecord 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   480
      Index           =   0
      Left            =   2880
      Picture         =   "Form21.frx":1010
      Stretch         =   -1  'True
      Top             =   4620
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image BancoImg 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   660
      Index           =   0
      Left            =   4200
      Picture         =   "Form21.frx":18DA
      Stretch         =   -1  'True
      Top             =   1992
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
   Begin VB.Shape S1 
      BackColor       =   &H80000001&
      BackStyle       =   1  'Opaque
      BorderStyle     =   3  'Dot
      Height          =   576
      Left            =   5016
      Top             =   4488
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Image Cbo 
      Height          =   315
      Index           =   0
      Left            =   1200
      Picture         =   "Form21.frx":2C24
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
      BackColor       =   &H80000004&
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
Dim RecX, RecY As Long, RecTipo As Boolean
Dim TmX, TmY As Long, TmTipo As Boolean
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
Dim Intervalor(0 To 1000) As Long

Private Sub BancoImg_Click(index As Integer)
On Error Resume Next
Set NovoObj = BancoImg(index)
Redimesiona
Selecione1
End Sub

Private Sub Cbo_DblClick(index As Integer)
Click2 Cbo(index).Tag
End Sub

Private Sub Chk_GotFocus(index As Integer)
Click2Var = 0
End Sub

Private Sub Chk_MouseDown(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
ChkY = Y
ChkX = X
ChkTipo = True
TesteBorda Button
End Sub


Private Sub Fm_DblClick(index As Integer)
Click2 Fm(index).Tag
End Sub

Private Sub Form_DblClick()
On Error Resume Next
Click2 ""
End Sub

Private Sub Form_GotFocus()
Form_Click
TelaAtual = IIf(Trim(Cont.Caption) = "", 0, Cont.Caption)
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Redimesiona
End Sub

Private Sub Img_DblClick(index As Integer)
Click2 Img(index).Tag
End Sub

Private Sub ImgRecord_Click(index As Integer)
Set NovoObj = ImgRecord(index)
Redimesiona
Selecione1
End Sub

Private Sub Lbl_DblClick(index As Integer)
Click2 Lbl(index).Tag
End Sub

Private Sub Lst_Click(index As Integer)
On Error Resume Next
FrmPrincipal.CboObj.Text = Lst(index).Tag
End Sub

Private Sub Lst_DblClick(index As Integer)
Click2 Lst(index).Tag
End Sub

Private Sub Lst_MouseMove(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If LstTipo And Me.MousePointer = 0 Then
    Lst(index).Move Lst(index).Left + X - LstX, Lst(index).Top + Y - LstY
    'Redimesiona
ElseIf Me.MousePointer = 2 And Button = 1 Then
    Form_MouseMove Button, Shift, X, Y
End If
End Sub

Private Sub Lst_MouseUp(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
LstTipo = False
If Me.MousePointer = 1 And Button = 1 Then
    Form_MouseDown Button, Shift, X, Y
End If
If NovoObj.Name <> "Lst" Then
    Set NovoObj = Lst(index)
End If
Selecione1
Redimesiona
End Sub
Private Sub Lst_MouseDown(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'If Button = 1 And Me.MousePointer = 1 Then
'    Set NovoObj = Lst(Index)
    '' Selecione1
    'Redimesiona
'ElseIf Me.MousePointer = 2 And Button = 1 Then
    'Form_MouseDown Button, Shift, X, Y
'End If
LstY = Y
LstX = X
LstTipo = True
TesteBorda Button
End Sub

Private Sub Chk_MouseMove(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If ChkTipo Then
    Chk(index).Move Chk(index).Left + X - ChkX, Chk(index).Top + Y - ChkY
    Click2Var = 0
    'Redimesiona
End If
End Sub

Private Sub Chk_MouseUp(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
ChkTipo = False
Redimesiona
Selecione1
End Sub

Private Sub Fm_Click(index As Integer)
On Error Resume Next
FrmPrincipal.CboObj.Text = Fm(index).Tag
Set NovoObj = Fm(index)
Redimesiona
Selecione1
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = vbKeyDelete Then
    If TypeOf NovoObj Is Form Then
        Exit Sub
    End If
    Dim Xp As Long

    For Xp = 1 To FrmCodigo.Eventos.Nodes.Count
        If UCase(FrmCodigo.Eventos.Nodes(Xp).Text) = UCase(NovoObj.Tag) Then
            FrmCodigo.Eventos.Nodes.Remove Xp
            Exit For
        End If
    Next Xp

    Im(0).Visible = False
    Im(1).Visible = False
    Im(2).Visible = False
    Im(3).Visible = False
    Im(4).Visible = False
    Im(5).Visible = False
    Im(6).Visible = False
    Im(7).Visible = False
    Unload NovoObj
    Form_Click
    If Err = 340 Then
        Exit Sub
    End If
    TesteBorda 1
    ' MsgBox
ElseIf KeyCode = vbKeyLeft And Shift = 0 Then
    NovoObj.Left = NovoObj.Left - 25
    'Redimesiona
    TesteBorda 1
ElseIf KeyCode = vbKeyRight And Shift = 0 Then
    NovoObj.Left = NovoObj.Left + 25
    'Redimesiona
    TesteBorda 1
ElseIf KeyCode = vbKeyUp And Shift = 0 Then
    NovoObj.Top = NovoObj.Top - 25
    'Redimesiona
    TesteBorda 1
ElseIf KeyCode = vbKeyDown And Shift = 0 Then
    NovoObj.Top = NovoObj.Top + 25
    'Redimesiona
    TesteBorda 1
ElseIf KeyCode = vbKeyLeft And Shift = 1 Then
    NovoObj.Width = NovoObj.Width - 25
    'Redimesiona
    TesteBorda 1
ElseIf KeyCode = vbKeyRight And Shift = 1 Then
    NovoObj.Width = NovoObj.Width + 25
    'Redimesiona
    TesteBorda 1
ElseIf KeyCode = vbKeyUp And Shift = 1 Then
    NovoObj.Height = NovoObj.Height - 25
    'Redimesiona
    TesteBorda 1
ElseIf KeyCode = vbKeyDown And Shift = 1 Then
    NovoObj.Height = NovoObj.Height + 25
    'Redimesiona
    TesteBorda 1
ElseIf Shift = 1 And KeyCode = 16 Then

Else
    FrmPrincipal.Grid.SetFocus
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

If FrmPrincipal.T.Buttons(1).Value = tbrPressed Then
    Form_Click
    Exit Sub
End If


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
ElseIf FrmPrincipal.T.Buttons(11).Value = tbrPressed Then
    CrieObj 9
    Me.MousePointer = 1
    FrmPrincipal.T.Buttons(11).Value = tbrUnpressed
ElseIf FrmPrincipal.T.Buttons(12).Value = tbrPressed Then
    CrieObj 10
    Me.MousePointer = 1
    FrmPrincipal.T.Buttons(11).Value = tbrUnpressed
End If
FrmPrincipal.T.Buttons(1).Value = tbrPressed
'Skin1.ApplySkin Me.HWnd
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'    ReadyToClose = True

 '   Cancel = Not ReadyToClose
End Sub

Private Sub Cbo_Click(index As Integer)
On Error Resume Next
FrmPrincipal.CboObj.Text = Cbo(index).Tag
Set NovoObj = Cbo(index)
Selecione1
Redimesiona
End Sub


Private Sub Chk_Click(index As Integer)
On Error Resume Next
FrmPrincipal.CboObj.Text = Chk(index).Tag
Chk(index).Value = 0
Set NovoObj = Chk(index)
IndexObj = index
Redimesiona
Selecione1
Click2Var = Click2Var + 1
If Click2Var = 2 Then
    'Click2 Chk(Index).Tag
End If
End Sub

Private Sub Cmd_Click(index As Integer)
On Error Resume Next
Click2Var = Click2Var + 1
If Click2Var = 2 Then
    Click2 Cmd(index).Tag
    FrmPrincipal.CboObj.Text = Cmd(index).Tag
    Set NovoObj = Cmd(index)
    Redimesiona
    Selecione1
    Exit Sub
End If
FrmPrincipal.CboObj.Text = Cmd(index).Tag
Set NovoObj = Cmd(index)
Redimesiona
Selecione1
Focus.SetFocus
End Sub

Private Sub Form_Activate()
On Error Resume Next
'Skin1.LoadSkin NomeSkin
'Skin1.ApplySkin Me.HWnd
If Cont.Caption <> "" Then
    Selecione1
'    Menu.Chama Menus(Cont.Caption)
'    Me.SetFocus
End If
Me.ScaleMode = 1

End Sub

Private Sub Im_Click(index As Integer)
'MoveAdd Index
End Sub

Private Sub Im_MouseDown(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
TipoBorda(index) = True
BordaX(index) = X
BordaY(index) = Y
TesteBorda 1
End Sub

Private Sub Im_MouseMove(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    TesteBorda 1
    On Error Resume Next
    If TipoBorda(index) Then
        If index = 0 Then
            Im(0).Move Im(0).Left + X - BordaX(index), Im(0).Top + Y - BordaY(index)
            NovoObj.Height = NovoObj.Height - (Im(0).Top - NovoObj.Top)
            NovoObj.Width = NovoObj.Width - (Im(0).Left - NovoObj.Left)
            NovoObj.Top = Im(0).Top
            NovoObj.Left = Im(0).Left
        ElseIf index = 1 Then
            Im(0).Move Im(0).Left, Im(0).Top + Y - BordaY(index)
            NovoObj.Height = NovoObj.Height - (Im(0).Top - NovoObj.Top)
        '    NovoObj.Width = NovoObj.Width - (Im(0).Left - NovoObj.Left)
            NovoObj.Top = Im(0).Top
         '   NovoObj.Left = Im(0).Left
        ElseIf index = 2 Then
            Im(2).Move Im(2).Left + X - BordaX(index), Im(2).Top + Y - BordaY(index)
            NovoObj.Height = NovoObj.Height - (Im(2).Top - NovoObj.Top)
            NovoObj.Top = Im(2).Top
            NovoObj.Width = Im(2).Left - Im(0).Left
            'NovoObj.Left = Im(2).Left - NovoObj.Width
        ElseIf index = 3 Then
            Im(3).Move Im(3).Left + X - BordaX(index), Im(3).Top + Y - BordaY(index)
            NovoObj.Width = NovoObj.Width + (NovoObj.Left - Im(3).Left)    'Im(3).Left + NovoObj.Left
            NovoObj.Left = Im(3).Left
            NovoObj.Height = Im(3).Top - NovoObj.Top
        ElseIf index = 4 Then
            Im(4).Move Im(4).Left + X - BordaX(index), Im(4).Top + Y - BordaY(index)
            NovoObj.Height = Im(4).Top - NovoObj.Top
        ElseIf index = 5 Then
            Im(5).Move Im(5).Left + X - BordaX(index), Im(5).Top + Y - BordaY(index)
            If (Im(5).Top - NovoObj.Top) > 0 Then
                NovoObj.Height = Im(5).Top - NovoObj.Top
            End If
            If (Im(5).Left - NovoObj.Left) > 0 Then
                NovoObj.Width = (Im(5).Left - NovoObj.Left)
            End If
        ElseIf index = 6 Then
            Im(6).Move Im(6).Left + X - BordaX(index), Im(6).Top + Y - BordaY(index)
            NovoObj.Width = NovoObj.Width - (Im(6).Left - NovoObj.Left)
            NovoObj.Left = Im(6).Left
        ElseIf index = 7 Then
            Im(7).Move Im(7).Left + X - BordaX(index), Im(7).Top + Y - BordaY(index)
            NovoObj.Width = Im(7).Left - NovoObj.Left
        End If
        'Redimesiona
        'Selecione1
    End If
End If
End Sub

Private Sub Im_MouseUp(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'TipoBorda(Index) = False
'If NovoObj.Name = "Form2" Then
'    Selecione 1, CCur(IndexObj)
'ElseIf NovoObj.Name = "Cmd" Then
'    Selecione 0, CCur(IndexObj)
'End If
'Im(0).Visible = True
'Im(1).Visible = True
'Im(2).Visible = True
'Im(3).Visible = True
'Im(4).Visible = True
'Im(5).Visible = True
'Im(6).Visible = True
'Im(7).Visible = True
Redimesiona
Selecione1
End Sub

Private Sub Img_Click(index As Integer)
On Error Resume Next
FrmPrincipal.CboObj.Text = Img(index).Tag
Set NovoObj = Img(index)
Selecione1
Redimesiona
End Sub

Private Sub Img_MouseDown(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
TesteBorda Button
ImgTipo = True
ImgX = X
ImgY = Y
End Sub
Private Sub BancoImg_MouseDown(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
BancoTipo = True
BancoX = X
BancoY = Y
TesteBorda 1
End Sub
Private Sub BancoImg_MouseMove(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If BancoTipo And Me.MousePointer = 0 Then
    BancoImg(index).Move BancoImg(index).Left + X - BancoX, BancoImg(index).Top + Y - BancoY
'    Redimesiona
End If
End Sub
Private Sub Img_MouseMove(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If ImgTipo And Me.MousePointer = 0 Then
    Img(index).Move Img(index).Left + X - ImgX, Img(index).Top + Y - ImgY
    'Redimesiona
End If
End Sub

Private Sub Img_MouseUp(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
ImgTipo = False
Selecione1
Redimesiona
End Sub
Private Sub BancoImg_MouseUp(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
BancoTipo = False
Redimesiona
End Sub

Private Sub Cmd_MouseDown(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
CmdTipo = True
CmdX = X
CmdY = Y
TesteBorda Button
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
    'If FrmPrincipal.Grid.Rows <> 7 Then
    '    Selecione 1, 0
    'End If
    On Error Resume Next
    FrmPrincipal.Grid.Row = 10
    FrmPrincipal.Grid.Col = 2
    FrmPrincipal.Grid.Text = FrmTela(TelaAtual).Height
    FrmPrincipal.Grid.Row = 11
    FrmPrincipal.Grid.Col = 2
    FrmPrincipal.Grid.Text = FrmTela(TelaAtual).Width
    'Selecione 1, 0
End If

End Sub
Private Sub Lbl_Click(index As Integer)
On Error Resume Next
'Lbl(Index).BorderStyle = 1
FrmPrincipal.CboObj.Text = Lbl(index).Tag
Set NovoObj = Lbl(index)
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
'ElseIf O.Caption = "8" Then
'    TotalObj.TotalOpt = TotalObj.TotalOpt + 1
ElseIf O.Caption = "8" Then
    TotalObj.TotalBanco = TotalObj.TotalBanco + 1
ElseIf O.Caption = "9" Then
    TotalObj.TotalRecord = TotalObj.TotalRecord + 1
ElseIf O.Caption = "10" Then
    TotalObj.TotalTime = TotalObj.TotalTime + 1
End If
End Sub



Private Sub Timer1_Timer()
Click2Var = 0
End Sub

Private Sub Tm_Click(index As Integer)
Set NovoObj = Tm(index)
Selecione1
Redimesiona
End Sub

Private Sub Txt_Click(index As Integer)
On Error Resume Next
FrmPrincipal.CboObj.Text = Txt(index).Tag
Set NovoObj = Txt(index)
Redimesiona
Selecione1
End Sub

Private Sub Txt_DblClick(index As Integer)
Click2 Txt(index).Tag
End Sub

Private Sub Txt_GotFocus(index As Integer)
Focus.SetFocus
End Sub
Private Sub Txt_MouseMove(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If TxtTipo And Me.MousePointer = 0 Then
    Txt(index).Move Txt(index).Left + X - TxtX, Txt(index).Top + Y - TxtY
    'Redimesiona
End If
End Sub

Private Sub Txt_MouseUp(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
TxtTipo = False
Redimesiona
Selecione1
End Sub
Private Sub Txt_MouseDown(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
TxtTipo = True
TxtX = X
TxtY = Y
TesteBorda Button
End Sub

Private Sub Cmd_MouseMove(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 And Me.MousePointer = 0 Then '
    If CmdTipo Then
        Cmd(index).Move Cmd(index).Left + X - CmdX, Cmd(index).Top + Y - CmdY
        'Click2Var = 0
        'Redimesiona
    End If
End If
End Sub

Private Sub Cmd_MouseUp(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
CmdTipo = False
Selecione1
Redimesiona
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
    Nome = FrmPrincipal.CboObj.Text
    FrmPrincipal.CboObj.Clear
    FrmPrincipal.CboObj.AddItem Me.Tag
    For Each Ob In Me
        If Ob.Name = "xMenu" Or Ob.Name = "Im" Or Ob.Name = "Focus" Or Ob.Name = "S" Or Ob.Name = "Nome" Or Ob.Name = "Cont" Then
            GoTo proximo
        'ElseIf Ob.Index = 0 Then
        Else
            FrmPrincipal.CboObj.AddItem Ob.Tag
        End If
proximo:
    Next
    FrmPrincipal.CboObj.Text = Me.Tag
End If
End Sub

Private Sub Form_Load()
'Banco.DatabaseName = LocalBancodeDados
TesteBorda 1
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
TotalObj.TotalTime = 1
Form_Resize
Me.Left = 0
Me.Top = 0
'Me.Visible = True
If CRun.OpenExe = 0 Then FrmPrincipal.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Cancel = 0 Then
    If FechaGeral = True Then
        Me.Visible = False
        Cancel = 1
    End If
End If
End Sub


Private Sub Fm_MouseDown(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
FmTipo = True
FmX = X
FmY = Y
TesteBorda Button
End Sub
Private Sub Fm_MouseMove(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If FmTipo And Me.MousePointer = 1 Then
    Fm(index).Move Fm(index).Left + X - FmX, Fm(index).Top + Y - FmY
    'Redimesiona
End If
End Sub

Private Sub Fm_MouseUp(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
FmTipo = False
If FrmPrincipal.T.Buttons(1).Value = tbrUnpressed Then
    If Me.MousePointer <> 0 Then
        Form_MouseUp Button, Shift, X, Y
    End If
End If
Selecione1
Redimesiona
End Sub


Private Sub Lbl_MouseDown(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
' Lbl(Index).BorderStyle = 1
LblTipo = True
LblX = X
LblY = Y
TesteBorda Button
End Sub
Private Sub Lbl_MouseMove(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If LblTipo Then 'And Me.MousePointer = 1 Then
    Lbl(index).Move Lbl(index).Left + X - LblX, Lbl(index).Top + Y - LblY
    'Redimesiona
End If
End Sub

Private Sub Lbl_MouseUp(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
LblTipo = False
Selecione1
Redimesiona
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


Private Sub cbo_MouseMove(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If CboTipo And Me.MousePointer = 0 Then
    Cbo(index).Move Cbo(index).Left + X - CboX, Cbo(index).Top + Y - CboY
    'Redimesiona
End If
End Sub

Private Sub Cbo_MouseUp(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
CboTipo = False
Selecione1
Redimesiona
End Sub
Private Sub Cbo_MouseDown(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
CboTipo = True
CboX = X
CboY = Y
TesteBorda Button
End Sub
Private Sub Form_MouseDown(index As Integer, Shift As Integer, X As Single, Y As Single)
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

Private Sub CrieObj(index As Byte)
On Error Resume Next
Dim X As Long

If index = 0 Then  'Se for um CommandBon
       
    For X = 0 To Cmd.Count
        Err.Number = 0
        Set NovoObj = Cmd(X)
        Test = NovoObj.Tag
        If Err.Number <> 0 Then
            Load Cmd(X)
            Set NovoObj = Cmd(X)
            Test = "Botao" + Trim(Str(X))
            For Xx = 1 To 1000
                Test = "Botao" + Trim(Str(Xx))
                For Yy = 0 To Cmd.Count - 1
                    If UCase(Test) = UCase(Cmd(Yy).Tag) Then
                        GoTo PCmd
                    End If
                Next Yy
                GoTo PPCmd
PCmd:
            Next Xx
PPCmd:
            NovoObj.Tag = Test
            GoTo Passa_Cmd
        End If
    Next X
   
Proximo_Cmd:
    Load Cmd(X)
    Set NovoObj = Cmd(X)
    NovoObj.Tag = "Botao" + Trim(Str(X))
Passa_Cmd:

ElseIf index = 1 Then 'Se for um Frame

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
ElseIf index = 2 Then 'Se for um Imagem
    
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

ElseIf index = 3 Then 'Se For um Label
    
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
    
ElseIf index = 4 Then 'Se for uma Check Box
    
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
ElseIf index = 5 Then 'Se for um Combo Box
    
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
ElseIf index = 6 Then

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
ElseIf index = 7 Then

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
ElseIf index = 8 Then
    Load BancoImg(TotalObj.TotalBanco)
    Set NovoObj = BancoImg(TotalObj.TotalBanco)
    BancoImg(TotalObj.TotalBanco).Tag = "Banco" + Trim(Str(TotalObj.TotalBanco))
    TotalObj.TotalBanco = TotalObj.TotalBanco + 1
    NovoObj.Width = 600
    NovoObj.Height = 600
ElseIf index = 9 Then
    Load ImgRecord(TotalObj.TotalRecord)
    Set NovoObj = ImgRecord(TotalObj.TotalRecord)
    TabRec(NovoObj.index).Nome = "Tabela" + Trim(Str(TotalObj.TotalRecord))
    TotalObj.TotalRecord = TotalObj.TotalRecord + 1
    NovoObj.Width = 480
    NovoObj.Height = 480
ElseIf index = 10 Then
    Load Tm(TotalObj.TotalTime)
    Set NovoObj = Tm(Tm.Count - 1)
    NovoObj.Tab = "Tempo" & TotalObj.TotalTime
    TotalObj.TotalTime = TotalObj.TotalTime + 1
    'set NovoObj
Else
    Exit Sub
End If

If index <> 8 And index <> 9 And index <> 10 Then
    NovoObj.Caption = NovoObj.Tag
    NovoObj.TabIndex = NovoObj.TabIndex - 19
    NovoObj.Height = S.Height
    NovoObj.Width = S.Width
End If
NovoObj.Left = S.Left
NovoObj.Top = S.Top
NovoObj.Visible = True
If index = 8 Or index = 9 Then
    Exit Sub
ElseIf index = 10 Then

Else

End If

FrmCodigo.Eventos.Nodes.Add UCase(Me.Tag), tvwChild, UCase(Me.Tag + "." + NovoObj.Tag), NovoObj.Tag, 3


With FrmCodigo.Eventos

    If index <> 8 And index <> 9 Then
        .Nodes.Add UCase(Me.Tag + "." + NovoObj.Tag), tvwChild, UCase(Me.Tag + "." + NovoObj.Tag + ".1"), "Ao Clicar 1 Vezes", 2
        If index = 7 Or index = 6 Or index = 5 Or index = 3 Or index = 2 Or index = 1 Then
            .Nodes.Add UCase(Me.Tag + "." + NovoObj.Tag), tvwChild, UCase(Me.Tag + "." + NovoObj.Tag + ".2"), "Ao Clicar 2 Vezes", 2
        End If
        .Nodes.Add UCase(Me.Tag + "." + NovoObj.Tag), tvwChild, UCase(Me.Tag + "." + NovoObj.Tag + ".Ganhar"), "Ao Ganhar o focu", 2
        .Nodes.Add UCase(Me.Tag + "." + NovoObj.Tag), tvwChild, UCase(Me.Tag + "." + NovoObj.Tag + ".Perder"), "Ao Perder o focu", 2
        If index = 6 Or 3 Then
            .Nodes.Add UCase(Me.Tag + "." + NovoObj.Tag), tvwChild, UCase(Me.Tag + "." + NovoObj.Tag + ".Escrever"), "Ao Escrever", 2
        End If
'    Else
'         .Nodes.Add UCase(Me.Tag + "." + NovoObj.Tag), tvwChild, UCase(Me.Tag + "." + NovoObj.Tag + ".Pular"), "Ao Mudar de Registro", 2
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

Public Sub TesteBorda(Button As Integer)
If Button = 1 Then
    Im(0).Visible = False
    Im(1).Visible = False
    Im(2).Visible = False
    Im(3).Visible = False
    Im(4).Visible = False
    Im(5).Visible = False
    Im(6).Visible = False
    Im(7).Visible = False
End If
End Sub
Private Sub ImgRecord_MouseDown(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
RecTipo = True
RecX = X
RecY = Y
TesteBorda 1
End Sub
Private Sub ImgRecord_MouseMove(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If RecTipo And Me.MousePointer = 0 Then
    ImgRecord(index).Move ImgRecord(index).Left + X - RecX, ImgRecord(index).Top + Y - RecY
    'Redimesiona
End If
End Sub
Private Sub ImgRecord_MouseUp(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
RecTipo = False
Redimesiona
End Sub


Private Sub Tm_MouseDown(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'If Button = 1 And Me.MousePointer = 1 Then
'    Set NovoObj = Lst(Index)
    '' Selecione1
    'Redimesiona
'ElseIf Me.MousePointer = 2 And Button = 1 Then
    'Form_MouseDown Button, Shift, X, Y
'End If
TmY = Y
TmX = X
TmTipo = True
End Sub

Private Sub Tm_MouseMove(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If TmTipo Then
    Tm(index).Move Tm(index).Left + X - TmX, Tm(index).Top + Y - TmY
End If
End Sub

Private Sub Tm_MouseUp(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
TmTipo = False
Redimesiona
Selecione1
End Sub
