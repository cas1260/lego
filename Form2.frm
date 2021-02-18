VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   4920
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6150
   Icon            =   "Form2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   NegotiateMenus  =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   6150
   Visible         =   0   'False
   Begin Projet1.OnFormMenu Menu 
      Height          =   345
      Left            =   690
      TabIndex        =   16
      Top             =   -90
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   609
   End
   Begin VB.TextBox TxtM 
      Height          =   1035
      Index           =   0
      Left            =   4500
      MousePointer    =   1  'Arrow
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   14
      Top             =   0
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.TextBox Focus 
      Height          =   285
      Left            =   1950
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   3150
      Width           =   1335
   End
   Begin VB.CheckBox Chk 
      Height          =   255
      Index           =   0
      Left            =   60
      TabIndex        =   5
      Top             =   60
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.CommandButton Cmd 
      Height          =   525
      Index           =   0
      Left            =   3810
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3390
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.TextBox Txt 
      Height          =   285
      Index           =   0
      Left            =   1440
      MousePointer    =   1  'Arrow
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.PictureBox Img 
      Height          =   525
      Index           =   0
      Left            =   2910
      ScaleHeight     =   465
      ScaleWidth      =   1185
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.Frame Fm 
      Height          =   525
      Index           =   0
      Left            =   3720
      TabIndex        =   3
      Top             =   2340
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.Label O 
      Height          =   285
      Left            =   4230
      TabIndex        =   15
      Top             =   4260
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.Shape S 
      Height          =   45
      Left            =   1140
      Top             =   2310
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Label Cl 
      Caption         =   "False"
      Height          =   225
      Left            =   60
      TabIndex        =   13
      Top             =   1920
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Label Ma 
      Caption         =   "False"
      Height          =   135
      Left            =   30
      TabIndex        =   12
      Top             =   1620
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Mi 
      Caption         =   "False"
      Height          =   165
      Left            =   60
      TabIndex        =   11
      Top             =   1260
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Label Max 
      Caption         =   "Não"
      Height          =   165
      Left            =   4380
      TabIndex        =   10
      Top             =   1170
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image Cbo 
      Height          =   315
      Index           =   0
      Left            =   1200
      Picture         =   "Form2.frx":0442
      Stretch         =   -1  'True
      Top             =   1080
      Visible         =   0   'False
      Width           =   2100
   End
   Begin VB.Image Mnu 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   510
      Left            =   3120
      Picture         =   "Form2.frx":26F8
      Top             =   3510
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Label P 
      Height          =   405
      Left            =   1230
      TabIndex        =   8
      Top             =   3600
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.Label Nome 
      Height          =   345
      Left            =   150
      TabIndex        =   7
      Top             =   4080
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Label Cont 
      Height          =   165
      Left            =   3600
      TabIndex        =   6
      Top             =   4560
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.Label Lbl 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   165
      Index           =   0
      Left            =   960
      TabIndex        =   1
      Top             =   1950
      Visible         =   0   'False
      Width           =   705
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
Private Sub Fm_Click(Index As Integer)
AcessoObj = "3"
IndexObj = Index
'Retaillable Fm(Index), True
FrmPrincipal.Banco.RecordSource = "Form" + Trim(Str(TelaAtual)) + "Fm" + Trim(Str(Index))
FrmPrincipal.Banco.Refresh
End Sub

Private Sub Focus_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = vbKeyDelete Then
    If AcessoObj = "1" Then
        Unload Txt(IndexObj)
    ElseIf AcessoObj = "2" Then
        Unload Lbl(IndexObj)
    ElseIf AcessoObj = "3" Then
        Unload Fm(IndexObj)
    ElseIf AcessoObj = "4" Then
        Unload Cbo(IndexObj)
    End If
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Me.MousePointer = 2 Then
    If FormTipo Then
        Me.Caption = x & "   " & Y
        If Y > S.Top Then
            S.Height = Y - FormY
        End If
        If x > S.Left Then
            S.Width = x - FormX
        End If
    End If
End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
FormTipo = False
S.Visible = False
If O.Caption = "1" Then
    Tool 1
    O.Caption = 0
    Me.MousePointer = 1
    FrmPrincipal.T.Buttons(2).Value = tbrUnpressed
ElseIf O.Caption = "4" Then
    Tool 4
    O.Caption = 0
    Me.MousePointer = 1
    FrmPrincipal.T.Buttons(5).Value = tbrUnpressed
ElseIf O.Caption = "6" Then
    Tool 6
    O.Caption = 0
    Me.MousePointer = 1
    FrmPrincipal.T.Buttons(7).Value = tbrUnpressed
ElseIf O.Caption = "2" Then
    Tool 2
    O.Caption = 0
    Me.MousePointer = 1
    FrmPrincipal.T.Buttons(4).Value = tbrUnpressed
End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'    ReadyToClose = True

 '   Cancel = Not ReadyToClose
End Sub

Private Sub Cbo_Click(Index As Integer)
Tamanho = Index
AcessoObj = "4"
AddGrind 4, Index
IndexObj = Index
P.Caption = "4"
End Sub


Private Sub Chk_Click(Index As Integer)
Chk(Index).Value = 0
End Sub

Private Sub Cmd_Click(Index As Integer)
Tamanho = Index
'S.Visible = True
'S.Top = Cmd(Index).Top - 50
'S.Left = Cmd(Index).Left - 50
'S.Height = Cmd(Index).Height + 85
'S.Width = Cmd(Index).Width + 85
Retaillable Cmd(Index), True
P.Caption = "2"
AddGrind 2, Index
IndexObj = Index
End Sub

Private Sub Cmd_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Then
    Unload Cmd(Index)
    S.Visible = False
    AddGrind 1, 0
End If
End Sub
Private Sub Lbl_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Then
    Unload Lbl(Index)
    AddGrind 1, 0
End If
End Sub


Private Sub Form_Activate()
On Error Resume Next
If Cont.Caption <> "" Then
    Menu.Chama Menus(Cont.Caption)
'    Me.SetFocus
End If
End Sub

Private Sub Mnu_DblClick()
On Error Resume Next
Dim List As ListBox

Set List = FrmMenu.ListProg

With FrmTela(TelaAtual)
    If Menus(.Cont.Caption).mp1.Visible = True Then
        x = 0
        List.AddItem Menus(.Cont.Caption).mp1.Caption
        Do While x <> Menus(.Cont.Caption).ma1.Count
            List.AddItem ">" + Menus(.Cont.Caption).ma1(x).Caption
            x = x + 1
        Loop
    End If
    If Menus(.Cont.Caption).mp2.Visible = True Then
        x = 0
        List.AddItem Menus(.Cont.Caption).mp2.Caption
        Do While x <> Menus(.Cont.Caption).ma2.Count
            List.AddItem ">" + Menus(.Cont.Caption).ma2(x).Caption
            x = x + 1
        Loop
    End If
    If Menus(.Cont.Caption).mp3.Visible = True Then
        x = 0
        List.AddItem Menus(.Cont.Caption).mp3.Caption
        Do While x <> Menus(.Cont.Caption).ma3.Count
            List.AddItem ">" + Menus(.Cont.Caption).ma3(x).Caption
            x = x + 1
        Loop
    End If
    If Menus(.Cont.Caption).mp4.Visible = True Then
        x = 0
        List.AddItem Menus(.Cont.Caption).mp4.Caption
        Do While x <> Menus(.Cont.Caption).ma4.Count
            List.AddItem ">" + Menus(.Cont.Caption).ma4(x).Caption
            x = x + 1
        Loop
    End If
    If Menus(.Cont.Caption).mp5.Visible = True Then
        x = 0
        List.AddItem Menus(.Cont.Caption).mp5.Caption
        Do While x <> Menus(.Cont.Caption).ma5.Count
            List.AddItem ">" + Menus(.Cont.Caption).ma5(x).Caption
            x = x + 1
        Loop
    End If
    If Menus(.Cont.Caption).mp6.Visible = True Then
        x = 0

        List.AddItem Menus(.Cont.Caption).mp6.Caption
        Do While x <> Menus(.Cont.Caption).ma6.Count
            List.AddItem ">" + Menus(.Cont.Caption).ma6(x).Caption
            x = x + 1
        Loop
    End If
    If Menus(.Cont.Caption).Mp7.Visible = True Then
        x = 0
        List.AddItem Menus(.Cont.Caption).Mp7.Caption
        Do While x <> Menus(.Cont.Caption).ma7.Count
            List.AddItem ">" + Menus(.Cont.Caption).ma7(x).Caption
            x = x + 1
        Loop

    End If
    If Menus(.Cont.Caption).mp8.Visible = True Then
        x = 0
        List.AddItem Menus(.Cont.Caption).mp8.Caption
        Do While x <> Menus(.Cont.Caption).ma8.Count
            List.AddItem ">" + Menus(.Cont.Caption).ma8(x).Caption
            x = x + 1
        Loop
    
    End If
End With
FrmMenu.Show 1
End Sub

Private Sub Txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Then
    Unload Txt(Index)
    AddGrind 1, 0
End If
End Sub
Private Sub Chk_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Then
    Unload Chk(Index)
    AddGrind 1, 0
End If
End Sub
Private Sub Fm_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Then
    Unload Fm(Index)
    AddGrind 1, 0
End If
End Sub
Private Sub Cbo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Then
    Unload Cbo(Index)
    AddGrind 1, 0
End If
End Sub


Private Sub Cmd_LostFocus(Index As Integer)
Retaillable Cmd(Index), False
FrmPrincipal.Grind.Col = 1
FrmPrincipal.Grind.Row = 3
FrmPrincipal.Grind.Columns(1) = Cmd(Index).Height
FrmPrincipal.Grind.Row = 4
FrmPrincipal.Grind.Columns(1) = Cmd(Index).Width
'S.Visible = False
End Sub

Private Sub Cmd_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
CmdTipo = True
CmdX = x
CmdY = Y
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
Menu.Top = 0
Menu.Left = 0
Menu.Width = Me.ScaleWidth
If TelaAtual <> -1 Then
'    If Left(UCase(FrmPrincipal.Banco.RecordSource), 4) <> "FORM" Then
'        FrmPrincipal.Banco.DatabaseName = NomeBanco
'        FrmPrincipal.Banco.RecordSource = "Form" + Trim(Str(TelaAtual))
'        FrmPrincipal.Banco.Refresh
'    End If
'    FrmPrincipal.Banco.Recordset.MoveFirst
'    FrmPrincipal.Banco.Recordset.Move 3
    FrmPrincipal.Grind.Row = 3
    FrmPrincipal.Grind.Columns(1) = Me.Height
    FrmPrincipal.Grind.Row = 4
    FrmPrincipal.Grind.Columns(1) = Me.Width
    
'    FrmPrincipal.Banco.Recordset.Edit
'    FrmPrincipal.Banco.Recordset!Campo2 = Me.Height
'    FrmPrincipal.Banco.Recordset.Update
'    FrmPrincipal.Banco.Recordset.MoveNext
'    FrmPrincipal.Banco.Recordset.Edit
'    FrmPrincipal.Banco.Recordset!Campo2 = Me.Width
'    FrmPrincipal.Banco.Recordset.Update
End If
'AddGrind 1, 0
End Sub


Private Sub Lbl_Click(Index As Integer)
'Lbl(Index).BorderStyle = 1
AcessoObj = "2"
FrmPrincipal.Banco.RecordSource = "Form" + Trim(Str(TelaAtual)) + "Lbl" + Trim(Str(Index))
FrmPrincipal.Banco.Refresh
'Retaillable Lbl(Index), True
P.Caption = "3"
AddGrind 3, Index
IndexObj = Index
End Sub


Private Sub Txt_Click(Index As Integer)
AcessoObj = "1"
FrmPrincipal.Banco.RecordSource = "Form" + Trim(Str(TelaAtual)) + "Txt" + Trim(Str(Index))
FrmPrincipal.Banco.Refresh
Retaillable Txt(Index), True
Tamanho = Index
P.Caption = "5"
AddGrind 5, Index
IndexObj = Index
End Sub

Private Sub Txt_GotFocus(Index As Integer)
Focus.SetFocus
End Sub

Private Sub Txt_LostFocus(Index As Integer)
Retaillable Txt(Index), False
End Sub

Private Sub Txt_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
If TxtTipo Then
    Txt(Index).Move Txt(Index).Left + x - TxtX, Txt(Index).Top + Y - TxtY
End If
End Sub

Private Sub Txt_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
TxtTipo = False
End Sub
Private Sub Txt_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
TxtTipo = True
TxtX = x
TxtY = Y
End Sub

Private Sub Cmd_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = 1 Then
    If CmdTipo Then
'        If InStr(1, FrmPrincipal.Banco.RecordSource, "Cmd") <> 0 Then
            Cmd(Index).Move Cmd(Index).Left + x - CmdX, Cmd(Index).Top + Y - CmdY
'        End If
        'S.Visible = True
        'S.Top = Cmd(Index).Top - 50
        'S.Left = Cmd(Index).Left - 50
        'S.Height = Cmd(Index).Height + 85
        'S.Width = Cmd(Index).Width + 85
    End If
End If
End Sub

Private Sub Cmd_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
CmdTipo = False
FrmPrincipal.Grind.Col = 1
FrmPrincipal.Grind.Row = 5
FrmPrincipal.Grind.Columns(1) = Cmd(Index).Top
FrmPrincipal.Grind.Row = 6
FrmPrincipal.Grind.Columns(1) = Cmd(Index).Left
End Sub


Private Sub Form_Click()
If P.Caption = "2" Then
    Retaillable Cmd(Tamanho), False
ElseIf P.Caption = "5" Then
    Retaillable Txt(Tamanho), False
ElseIf P.Caption = "6" Then
    Retaillable Fm(Tamanho), False
End If
S.Visible = False
TelaAtual = Cont.Caption
AddGrind 1, 0
End Sub

Private Sub Form_Load()
Form_Resize
Me.Left = 0
Me.Top = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Cancel = 0 Then
    Me.Visible = False
    Cancel = 1
End If
End Sub


Private Sub Fm_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
FmTipo = True
FmX = x
FmY = Y
End Sub
Private Sub Fm_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
If FmTipo Then
    Fm(Index).Move Fm(Index).Left + x - FmX, Fm(Index).Top + Y - FmY
End If
End Sub

Private Sub Fm_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
FmTipo = False
End Sub


Private Sub Lbl_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
Lbl(Index).BorderStyle = 1
LblTipo = True
LblX = x
LblY = Y
End Sub
Private Sub Lbl_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
If LblTipo Then
    Lbl(Index).Move Lbl(Index).Left + x - LblX, Lbl(Index).Top + Y - LblY
End If
End Sub

Private Sub Lbl_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
LblTipo = False
Lbl(Index).BorderStyle = 0
End Sub
Public Function Rest()
On Error Resume Next
Dim x As Long
x = 1
Do While x <> Cmd.Count
    Retaillable Cmd(x), False
    x = x + 1
Loop

End Function

Public Function AddGrind(Index As Long, Index2 As Integer)
On Error Resume Next
Select Case Index
    Case 1
    P.Caption = "1"
    With FrmPrincipal
        .Banco.RecordSource = "Form" + Trim(Str(TelaAtual))
        .Banco.Refresh
    End With
    Case 2
        FrmPrincipal.Banco.RecordSource = "Form" + Trim(Str(TelaAtual)) + "Cmd" + Trim(Str(Index2))
        FrmPrincipal.Banco.Refresh
End Select
End Function

Private Sub Mnu_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
If MenuTipo Then
    Mnu.Move Mnu.Left + x - MenuX, Mnu.Top + Y - MenuY
End If
End Sub

Private Sub Mnu_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
MenuTipo = False
End Sub
Private Sub Mnu_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
MenuTipo = True
MenuX = x
MenuY = Y
End Sub


Private Sub cbo_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
If CboTipo Then
    Cbo(Index).Move Cbo(Index).Left + x - CboX, Cbo(Index).Top + Y - CboY
    'Cbo(Index).Left = Cbo(Index).Left + X - CboX
    'Cbo(Index).Top = Cbo(Index).Top + Y - CboY
'    CboTipo = False
End If
End Sub

Private Sub Cbo_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
CboTipo = False
End Sub
Private Sub Cbo_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
CboTipo = True
CboX = x
CboY = Y
End Sub
Public Function Tool(Index As Integer)
On Error Resume Next
Dim Ut As Long, PX As Long, PY As Long
Dim Na() As String, IX As Long
If TelaAtual <> -1 Then
    PX = Me.ScaleHeight / 2
    PY = Me.ScaleWidth / 2
    Select Case Index
        Case Is = 1
            Ut = Cmd.Count
            Load Cmd(Ut)
            Cmd(Ut).ToolTipText = "Botao" + Trim(Str(Ut))
            Cmd(Ut).Top = PX
            Cmd(Ut).Left = PY
            Cmd(Ut).Visible = True
            Field "Form" + Trim(Str(TelaAtual)) + "Cmd" + Trim(Str(Ut)), 2
            ReDim Na(8) As String
            Cmd(Ut).Height = S.Height
            Cmd(Ut).Width = S.Width
            Cmd(Ut).Left = S.Left
            Cmd(Ut).Top = S.Top
            Na(0) = Cmd(Ut).ToolTipText
            Na(1) = Cmd(Ut).BackColor
            Na(2) = Cmd(Ut).Caption
            Na(3) = Cmd(Ut).Height
            Na(4) = Cmd(Ut).Width
            Na(5) = Cmd(Ut).Top
            Na(6) = Cmd(Ut).Left
            Na(7) = Cmd(Ut).FontName
            FrmPrincipal.Banco.RecordSource = "Form" + Trim(Str(TelaAtual)) + "Cmd" + Trim(Str(Ut))
            FrmPrincipal.Banco.Refresh
            FrmPrincipal.Banco.Recordset.MoveFirst
            With FrmPrincipal.Banco.Recordset
                For IX = 0 To 7
                    .Edit
                    !Campo2 = IIf(Trim(Na(IX)) = "", " ", Na(IX))
                    .Update
                    .MoveNext
                Next IX
            End With
        Case Is = 2
            Ut = Fm.Count
            Load Fm(Ut)
            Fm(Ut).ToolTipText = "Frame" + Trim(Str(Ut))
            Fm(Ut).Top = S.Top
            Fm(Ut).Left = S.Left
            Fm(Ut).Height = S.Height
            Fm(Ut).Width = S.Width
            Fm(Ut).Visible = True
            Field "Form" + Trim(Str(TelaAtual)) + "Fm" + Trim(Str(Ut)), 5
            
            ReDim Na(8) As String
            Na(0) = Fm(Ut).ToolTipText
            Na(1) = Fm(Ut).BackColor
            Na(2) = Fm(Ut).Caption
            Na(3) = Fm(Ut).Height
            Na(4) = Fm(Ut).Width
            Na(5) = Fm(Ut).Top
            Na(6) = Fm(Ut).Left
            Na(7) = Fm(Ut).FontName
            Na(8) = Fm(Ut).FontSize
            FrmPrincipal.Banco.RecordSource = "Form" + Trim(Str(TelaAtual)) + "Fm" + Trim(Str(Ut))
            FrmPrincipal.Banco.Refresh
            FrmPrincipal.Banco.Recordset.MoveFirst
            With FrmPrincipal.Banco.Recordset
                For IX = 0 To 8
                    .Edit
                    !Campo2 = IIf(Trim(Na(IX)) = "", " ", Na(IX))
                    .Update
                    .MoveNext
                Next IX
            End With
        Case Is = 3
            Ut = Img.Count
            Load Img(Ut)
            Img(Ut).Top = PX
            Img(Ut).Left = PY
            Img(Ut).Visible = True
        Case Is = 4
            Ut = Lbl.Count
            Load Lbl(Ut)
            Lbl(Ut).Caption = "Legenda " + Str(Ut)
            Lbl(Ut).ToolTipText = "Lgd" + Trim(Str(Ut))
            Lbl(Ut).Top = S.Top
            Lbl(Ut).Left = S.Left
            Lbl(Ut).Height = S.Height
            Lbl(Ut).Width = S.Width

            Lbl(Ut).Visible = True
            
            Field "Form" + Trim(Str(TelaAtual)) + "Lbl" + Trim(Str(Ut)), 3
                        
            ReDim Na(9) As String
            Na(0) = Lbl(Ut).ToolTipText
            Na(1) = Lbl(Ut).BackColor
            Na(2) = Lbl(Ut).ForeColor
            Na(3) = Lbl(Ut).Caption
            Na(4) = Lbl(Ut).Height
            Na(5) = Lbl(Ut).Width
            Na(6) = Lbl(Ut).Top
            Na(7) = Lbl(Ut).Left
            Na(8) = Lbl(Ut).FontName
            Na(9) = Lbl(Ut).FontSize
            
            FrmPrincipal.Banco.RecordSource = "Form" + Trim(Str(TelaAtual)) + "Lbl" + Trim(Str(Ut))
            FrmPrincipal.Banco.Refresh
            FrmPrincipal.Banco.Recordset.MoveFirst
            With FrmPrincipal.Banco.Recordset
                For IX = 0 To 9
                    .Edit
                    !Campo2 = IIf(Trim(Na(IX)) = "", " ", Na(IX))
                    .Update
                    .MoveNext
                Next IX
            End With
        Case Is = 5
            Ut = Chk.Count
            Load Chk(Ut)
            Chk(Ut).Top = PX
            Chk(Ut).Left = PY
            Chk(Ut).Visible = True
        Case Is = 6
            Ut = Txt.Count
            Field "Form" + Trim(Str(TelaAtual)) + "Txt" + Trim(Str(Ut)), 4
            Load Txt(Ut)
            Txt(Ut).ToolTipText = "Texto" + Trim(Str(Ut))
            Txt(Ut).Top = S.Top
            Txt(Ut).Left = S.Left
            Txt(Ut).Height = S.Height
            Txt(Ut).Width = S.Width
            Txt(Ut).Visible = True
            ReDim Na(9) As String
            Na(0) = Txt(Ut).ToolTipText
            Na(1) = Txt(Ut).BackColor
            Na(2) = Txt(Ut).ForeColor
            Na(3) = Txt(Ut).Text
            Na(4) = Txt(Ut).Height
            Na(5) = Txt(Ut).Width
            Na(6) = Txt(Ut).Top
            Na(7) = Txt(Ut).Left
            Na(8) = Txt(Ut).FontName
            Na(9) = Txt(Ut).FontSize
            
            FrmPrincipal.Banco.RecordSource = "Form" + Trim(Str(TelaAtual)) + "Txt" + Trim(Str(Ut))
            FrmPrincipal.Banco.Refresh
            FrmPrincipal.Banco.Recordset.MoveFirst
            With FrmPrincipal.Banco.Recordset
                For IX = 0 To 9
                    .Edit
                    !Campo2 = IIf(Trim(Na(IX)) = "", " ", Na(IX))
                    .Update
                    .MoveNext
                Next IX
            End With
        
        Case Is = 7
            Ut = Cbo.Count
            Load Cbo(Ut)
            Cbo(Ut).ToolTipText = "Combo" + Trim(Str(Ut))
            Cbo(Ut).Top = PX
            Cbo(Ut).Left = PY
            Cbo(Ut).Visible = True
        Case Is = 8
            Mnu.Visible = True
      End Select
End If

End Function
Private Sub Form_MouseDown(Index As Integer, Shift As Integer, x As Single, Y As Single)
S.Visible = True
S.Height = 0
S.Width = 0
S.Left = x
S.Top = Y
FormTipo = True
FormX = x
FormY = Y
'S.Visible = True
End Sub

