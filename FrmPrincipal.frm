VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{32A4927E-FB95-11D1-BF5B-00A024982E5B}#94.0#0"; "AXGRID.OCX"
Begin VB.MDIForm FrmPrincipal 
   BackColor       =   &H00808000&
   Caption         =   "Lego 1.0"
   ClientHeight    =   4830
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6030
   Icon            =   "FrmPrincipal.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2400
      Top             =   3420
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   28
      ImageHeight     =   28
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrincipal.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrincipal.frx":0DC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrincipal.frx":174A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrincipal.frx":20CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrincipal.frx":2A52
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrincipal.frx":33D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrincipal.frx":3D5A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrincipal.frx":46DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrincipal.frx":4E42
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrincipal.frx":57C6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Im2 
      Align           =   4  'Align Right
      Height          =   4830
      Left            =   3600
      ScaleHeight     =   4770
      ScaleWidth      =   2370
      TabIndex        =   1
      Top             =   0
      Width           =   2430
      Begin VB.ListBox Prog 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1080
         Left            =   270
         TabIndex        =   3
         Top             =   510
         Width           =   885
      End
      Begin axGridControl.axgrid Grind 
         Height          =   2355
         Left            =   150
         TabIndex        =   5
         Top             =   1980
         Width           =   2235
         _ExtentX        =   3942
         _ExtentY        =   4154
         Rows            =   0
         Cols            =   2
         Redraw          =   -1  'True
         ShowGrid        =   -1  'True
         GridSolid       =   0   'False
         GridLineColor   =   12632256
         BorderStyle     =   4
         BackColorFixed  =   16777215
         ColHeader       =   0   'False
         RowHeader       =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColorBkg    =   14737632
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000005&
         X1              =   0
         X2              =   2370
         Y1              =   330
         Y2              =   330
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   2340
         Y1              =   300
         Y2              =   300
      End
      Begin VB.Label L 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Projetos"
         Height          =   195
         Left            =   885
         TabIndex        =   4
         Top             =   60
         Width           =   585
      End
      Begin VB.Shape F2 
         BorderColor     =   &H80000005&
         Height          =   4815
         Left            =   -30
         Top             =   0
         Width           =   2415
      End
   End
   Begin VB.PictureBox Im 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H80000008&
      Height          =   4830
      Left            =   0
      ScaleHeight     =   4800
      ScaleWidth      =   1470
      TabIndex        =   0
      Top             =   0
      Width           =   1500
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   780
         Top             =   3930
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   2640
         Visible         =   0   'False
         Width           =   1140
      End
      Begin MSComctlLib.Toolbar T 
         Height          =   2610
         Left            =   60
         TabIndex        =   2
         Top             =   120
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   4604
         ButtonWidth     =   926
         ButtonHeight    =   900
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   9
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "A"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "B"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "C"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "D"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "E"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "F"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "G"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "H"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "J"
               ImageIndex      =   10
            EndProperty
         EndProperty
      End
      Begin VB.Shape F 
         BorderColor     =   &H80000005&
         FillColor       =   &H00FFFFFF&
         Height          =   4815
         Left            =   0
         Top             =   0
         Width           =   1485
      End
   End
   Begin VB.Menu MenuArquivo 
      Caption         =   "&Arquivo"
      Begin VB.Menu menuNovo 
         Caption         =   "&Novo"
         Begin VB.Menu MenuProjeto 
            Caption         =   "Projeto"
         End
         Begin VB.Menu MenuBranco 
            Caption         =   "-"
            Index           =   0
         End
         Begin VB.Menu MenuTela 
            Caption         =   "Tela"
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
      End
      Begin VB.Menu MenuBranco4 
         Caption         =   "-"
      End
      Begin VB.Menu menuComplile 
         Caption         =   "&Complilar             "
         Shortcut        =   {F5}
      End
      Begin VB.Menu haskfd 
         Caption         =   "-"
      End
      Begin VB.Menu MenuSair 
         Caption         =   "&Sair"
         Shortcut        =   ^Q
      End
   End
End
Attribute VB_Name = "FrmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim A As String
Dim I As Long
Private Sub Grind_AfterEdit(Row As Integer, Col As Integer, NewValue As String)
EditGrind Col, Row
End Sub

Private Sub Grind_BeforeEdit(Row As Integer, Col As Integer, ByVal Cancel As Boolean)
'If FrmTela(TelaAtual).P.Caption = "1" Then
   Grind.Text = ""
'End If

End Sub

Private Sub MDIForm_Load()
ReDim FrmTela(999) As New Form2
ReDim Run(999) As New Compl
ReDim Menus(999) As New FrmM
MDIForm_Resize
ContTela = 0
TelaAtual = -1
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
End
End Sub

Private Sub MDIForm_Resize()
If Me.WindowState <> 1 Then
    F.Top = 0
    F.Left = 0
    F.Height = Me.ScaleHeight
    F.Width = Im.Width
    F2.Top = 0
    F2.Left = 0
    F2.Height = Me.ScaleHeight
    F2.Width = Im2.Width
    T.Top = 100
    T.Left = 10
    T.Height = F.Height - 10
    T.Width = F.Width - 10
    Prog.Top = L.Height + 130
    Prog.Left = 10
    Prog.Height = (F2.Height / 2)
    Prog.Width = F2.Width - 80
    Grind.Top = Prog.Top + Prog.Height + 100
    Grind.Left = 0
    Grind.Height = F2.Height - (Prog.Height + 1000)
    Grind.Width = F2.Width
End If
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
End
End Sub

Private Sub menuComplile_Click()
RunProg
End Sub

Private Sub MenuSair_Click()
End
End Sub

Private Sub MenuSalvar_Click()
SalvarProjeto True
End Sub

Private Sub MenuTela_Click()
A = ""
A = InputBox("Nome da Tela :", App.Title)
If A = "" Then
    Exit Sub
End If

FrmTela(ContTela).Nome.Caption = A
FrmTela(ContTela).Cont = ContTela
FrmTela(ContTela).Caption = "Tela " + Str(ContTela)
Prog.AddItem A
FrmTela(ContTela).Show vbModeless
If ContTela = 0 Then
    TelaAtual = 0
End If
ContTela = ContTela + 1
End Sub
Private Sub Prog_DblClick()
If Prog.ListIndex > -1 Then
    FrmTela(Prog.ListIndex).Show vbModeless
End If
End Sub

Private Sub T_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
    Case Is = "B"
        Cmd 1
    Case Is = "C"
        Cmd 2
    Case Is = "D"
        Cmd 3
    Case Is = "E"
        Cmd 4
    Case Is = "F"
        Cmd 5
    Case Is = "G"
        Cmd 6
    Case Is = "H"
        Cmd 7
    Case Is = "J"
        Cmd 8
    End Select
End Sub

Public Function Cmd(Index As Integer)
On Error Resume Next
Dim Ut As Long, PX As Long, PY As Long
If TelaAtual <> -1 Then
    PX = FrmTela(TelaAtual).ScaleHeight / 2
    PY = FrmTela(TelaAtual).ScaleWidth / 2
    Select Case Index
        Case Is = 1
            Ut = FrmTela(TelaAtual).Cmd.Count
            Load FrmTela(TelaAtual).Cmd(Ut)
            FrmTela(TelaAtual).Cmd(Ut).ToolTipText = "Botao" + Trim(Str(Ut))
            FrmTela(TelaAtual).Cmd(Ut).Top = PX
            FrmTela(TelaAtual).Cmd(Ut).Left = PY
            FrmTela(TelaAtual).Cmd(Ut).Visible = True
        Case Is = 2
            Ut = FrmTela(TelaAtual).Fm.Count
            Load FrmTela(TelaAtual).Fm(Ut)
            FrmTela(TelaAtual).Fm(Ut).ToolTipText = "Frame" + Trim(Str(Ut))
            FrmTela(TelaAtual).Fm(Ut).Top = PX
            FrmTela(TelaAtual).Fm(Ut).Left = PY
            FrmTela(TelaAtual).Fm(Ut).Visible = True
        Case Is = 3
            Ut = FrmTela(TelaAtual).Img.Count
            Load FrmTela(TelaAtual).Img(Ut)
            FrmTela(TelaAtual).Img(Ut).Top = PX
            FrmTela(TelaAtual).Img(Ut).Left = PY
            FrmTela(TelaAtual).Img(Ut).Visible = True
        Case Is = 4
            Ut = FrmTela(TelaAtual).Lbl.Count
            Load FrmTela(TelaAtual).Lbl(Ut)
            FrmTela(telaautal).Lbl(Ut).Caption = "Legenda " + Str(Ut)
            FrmTela(TelaAtual).Lbl(Ut).ToolTipText = "Lgd" + Trim(Str(Ut))
            FrmTela(TelaAtual).Lbl(Ut).Top = PX
            FrmTela(TelaAtual).Lbl(Ut).Left = PY
            FrmTela(TelaAtual).Lbl(Ut).Visible = True
        Case Is = 5
            Ut = FrmTela(TelaAtual).Chk.Count
            Load FrmTela(TelaAtual).Chk(Ut)
            FrmTela(TelaAtual).Chk(Ut).Top = PX
            FrmTela(TelaAtual).Chk(Ut).Left = PY
            FrmTela(TelaAtual).Chk(Ut).Visible = True
        Case Is = 6
            Ut = FrmTela(TelaAtual).Txt.Count
            Load FrmTela(TelaAtual).Txt(Ut)
            FrmTela(TelaAtual).Txt(Ut).ToolTipText = "Texto" + Trim(Str(Ut))
            FrmTela(TelaAtual).Txt(Ut).Top = PX
            FrmTela(TelaAtual).Txt(Ut).Left = PY
            FrmTela(TelaAtual).Txt(Ut).Visible = True
        Case Is = 7
            Ut = FrmTela(TelaAtual).Cbo.Count
            Load FrmTela(TelaAtual).Cbo(Ut)
            FrmTela(TelaAtual).Cbo(Ut).ToolTipText = "Combo" + Trim(Str(Ut))
            FrmTela(TelaAtual).Cbo(Ut).Top = PX
            FrmTela(TelaAtual).Cbo(Ut).Left = PY
            FrmTela(TelaAtual).Cbo(Ut).Visible = True
        Case Is = 8
            FrmTela(TelaAtual).Mnu.Visible = True
      End Select
End If

End Function

Public Function EditGrind(Col As Integer, Row As Integer)
Dim Texto As String
If FrmTela(TelaAtual).P.Caption = "1" Then
    If Row = 1 Then
        For I = 0 To Prog.ListCount
            If UCase(FrmTela(TelaAtual).Nome.Caption) = UCase(Prog.List(I)) Then
                Prog.List(I) = Grind.Text
                FrmTela(TelaAtual).Nome.Caption = Grind.Text
            End If
        Next I
    ElseIf Row = 2 Then
        FrmTela(TelaAtual).BackColor = Left(Grind.Text, 7)
    ElseIf Row = 3 Then
        FrmTela(TelaAtual).Caption = Grind.Text
    ElseIf Row = 4 Then
        If IsNumeric(Grind.Text) = True Then
            FrmTela(TelaAtual).Height = Grind.Text
        End If
    ElseIf Row = 5 Then
        If IsNumeric(Grind.Text) = True Then
            FrmTela(TelaAtual).Width = Grind.Text
        End If
    ElseIf Row = 6 Then
        If UCase(Grind.Text) = "SIM" Then
            Grind.Text = "Sim"
            FrmTela(TelaAtual).Max.Caption = Grind.Text
        Else
            FrmTela(TelaAtual).Max.Caption = "Não"
            Grind.Text = "Não"
        End If
        Grind.Refresh
    End If
ElseIf FrmTela(TelaAtual).P.Caption = "2" Then
    If Row = 1 Then
        If Trim(Grind.Text) <> "" Then
            FrmTela(TelaAtual).Cmd(IndexObj).ToolTipText = Grind.Text
        Else
            Grind.Text = FrmTela(TelaAtual).Cmd(IndexObj).ToolTipText
        End If
    ElseIf Row = 2 Then
        FrmTela(TelaAtual).Cmd(IndexObj).Caption = Grind.Text
    ElseIf Row = 3 Then
        FrmTela(TelaAtual).Cmd(IndexObj).Height = Grind.Text
    ElseIf Row = 4 Then
        FrmTela(TelaAtual).Cmd(IndexObj).Width = Grind.Text
    ElseIf Row = 6 Then
        FrmTela(TelaAtual).Cmd(IndexObj).Top = Grind.Text
    ElseIf Row = 7 Then
        FrmTela(TelaAtual).Cmd(IndexObj).Left = Grind.Text
    End If
ElseIf FrmTela(TelaAtual).P.Caption = "3" Then
    With FrmTela(TelaAtual).Lbl(IndexObj)
        If Row = 1 Then
           .ToolTipText = Grind.Text
        ElseIf Row = 2 Then
           .Caption = Grind.Text
        ElseIf Row = 3 Then
           .Height = Grind.Text
        ElseIf Row = 4 Then
            .Width = Grind.Text
        ElseIf Row = 5 Then
            .Left = Grind.Text
        ElseIf Row = 6 Then
            .Top = Grind.Text
        ElseIf Row = 7 Then
            .BackColor = Left(Grind.Text, 8)
        ElseIf Row = 8 Then
            .ForeColor = Left(Grind.Text, 8)
        ElseIf Row = 9 Then
            .FontSize = Grind.Text
        ElseIf Row = 10 Then
            .FontName = Grind.Text
        End If
    End With
ElseIf FrmTela(TelaAtual).P.Caption = "4" Then
    With FrmTela(TelaAtual).Cbo(IndexObj)
        If Row = 1 Then
           .ToolTipText = Grind.Text
        ElseIf Row = 2 Then
            .Width = Grind.Text
        ElseIf Row = 5 Then
            .Left = Grind.Text
        ElseIf Row = 6 Then
            .Top = Grind.Text
        End If
    End With
        
ElseIf FrmTela(TelaAtual).P.Caption = "5" Then
    With FrmTela(TelaAtual).Txt(IndexObj)
        Select Case Row
            Case 1
                .ToolTipText = Grind.Text
            Case 2
                .Text = Grind.Text
            Case 3
                .BackColor = Grind.Text
            Case 4
                .ForeColor = Grind.Text
            Case 5
                .FontSize = Grind.Text
            Case 6
                .FontName = Grind.Text
            Case 7
                .Top = Grind.Text
            Case 8
                .Left = Grind.Text
            Case 9
                .Height = Grind.Text
            Case 10
                .Width = Grind.Text
            Case 11
                Texto = .Text
                .PasswordChar = Left(Grind.Text, 1)
                Grind.Text = Left(Grind.Text, 1)
                .Text = Texto
        End Select
    End With
End If

End Function
Function treste()
If FrmTela(TelaAtual).P.Caption = "1" Then
    If Row = 1 Then
        For I = 0 To Prog.ListCount
            If UCase(FrmTela(TelaAtual).Nome.Caption) = UCase(Prog.List(I)) Then
                Prog.List(I) = Grind.Text
                FrmTela(TelaAtual).Nome.Caption = Grind.Text
            End If
        Next I
    ElseIf Row = 2 Then
        FrmTela(TelaAtual).BackColor = Left(Grind.Text, 8)
    ElseIf Row = 3 Then
        FrmTela(TelaAtual).Caption = Grind.Text
    ElseIf Row = 4 Then
        If IsNumeric(Grind.Text) = True Then
            FrmTela(TelaAtual).Height = Grind.Text
        End If
    ElseIf Row = 5 Then
        If IsNumeric(Grind.Text) = True Then
            FrmTela(TelaAtual).Width = Grind.Text
        End If
    End If
ElseIf FrmTela(TelaAtual).P.Caption = "2" Then
    'MsgBox FrmTela(TelaAtual).Cmd(IndexObj).ToolTipText
    If Row = 1 Then
        If Trim(Grind.Text) <> "" Then
            FrmTela(TelaAtual).Cmd(IndexObj).ToolTipText = Grind.Text
        Else
            Grind.Text = FrmTela(TelaAtual).Cmd(IndexObj).ToolTipText
        End If
    ElseIf Row = 2 Then
        FrmTela(TelaAtual).Cmd(IndexObj).Caption = Grind.Text
    ElseIf Row = 3 Then
        FrmTela(TelaAtual).Cmd(IndexObj).Height = Grind.Text
    ElseIf Row = 4 Then
        FrmTela(TelaAtual).Cmd(IndexObj).Width = Grind.Text
    ElseIf Row = 6 Then
        FrmTela(TelaAtual).Cmd(IndexObj).Top = Grind.Text
    ElseIf Row = 7 Then
        FrmTela(TelaAtual).Cmd(IndexObj).Left = Grind.Text
    End If
ElseIf FrmTela(TelaAtual).P.Caption = "3" Then
    With FrmTela(TelaAtual).Lbl(IndexObj)
        If Row = 1 Then
           .ToolTipText = Grind.Text
        ElseIf Row = 2 Then
           .Caption = Grind.Text
        ElseIf Row = 3 Then
           .Height = Grind.Text
        ElseIf Row = 4 Then
            .Width = Grind.Text
        ElseIf Row = 5 Then
            .Left = Grind.Text
        ElseIf Row = 6 Then
            .Top = Grind.Text
        ElseIf Row = 7 Then
            .BackColor = Left(Grind.Text, 8)
        ElseIf Row = 8 Then
            .ForeColor = Left(Grind.Text, 8)
        ElseIf Row = 9 Then
            .FontSize = Grind.Text
        ElseIf Row = 10 Then
            .FontName = Grind.Text
        End If
    End With
ElseIf FrmTela(TelaAtual).P.Caption = "4" Then
    With FrmTela(TelaAtual).Cbo(IndexObj)
        If Row = 1 Then
           .ToolTipText = Grind.Text
        ElseIf Row = 2 Then
            .Width = Grind.Text
        ElseIf Row = 5 Then
            .Left = Grind.Text
        ElseIf Row = 6 Then
            .Top = Grind.Text
        End If
    End With
End If

End Function
