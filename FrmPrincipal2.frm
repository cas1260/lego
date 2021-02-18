VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.MDIForm FrmPrincipal 
   BackColor       =   &H00808000&
   Caption         =   "Lego 1.0"
   ClientHeight    =   4830
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6030
   Icon            =   "FrmPrincipal2.frx":0000
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
            Picture         =   "FrmPrincipal2.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrincipal2.frx":0DC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrincipal2.frx":174A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrincipal2.frx":20CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrincipal2.frx":2A52
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrincipal2.frx":33D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrincipal2.frx":3D5A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrincipal2.frx":46DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrincipal2.frx":4E42
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrincipal2.frx":57C6
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
      Begin MSDBGrid.DBGrid Grind 
         Bindings        =   "FrmPrincipal2.frx":5C1A
         Height          =   2235
         Left            =   180
         OleObjectBlob   =   "FrmPrincipal2.frx":5C2E
         TabIndex        =   5
         Top             =   2100
         Width           =   2145
      End
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
         Width           =   2475
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
      Begin VB.Data Banco 
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
         Height          =   570
         Left            =   60
         TabIndex        =   2
         Top             =   120
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   1005
         ButtonWidth     =   926
         ButtonHeight    =   900
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   9
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "A"
               ImageIndex      =   1
               Style           =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "B"
               ImageIndex      =   2
               Style           =   1
               Object.Width           =   1e-4
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "C"
               ImageIndex      =   3
               Style           =   1
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "D"
               ImageIndex      =   4
               Style           =   1
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "E"
               ImageIndex      =   5
               Style           =   1
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "F"
               ImageIndex      =   6
               Style           =   1
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "G"
               ImageIndex      =   7
               Style           =   1
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "H"
               ImageIndex      =   9
               Style           =   1
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "J"
               ImageIndex      =   10
               Style           =   1
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
         Caption         =   "&Executar          "
         Shortcut        =   {F5}
      End
      Begin VB.Menu meudddd 
         Caption         =   "-"
      End
      Begin VB.Menu MenuCompli 
         Caption         =   "&Compile                                      "
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
      Begin VB.Menu MenuPropriedade 
         Caption         =   "Propriedade"
         Shortcut        =   {F4}
      End
      Begin VB.Menu MenuBr 
         Caption         =   "-"
      End
      Begin VB.Menu MenuTelas 
         Caption         =   "Telas"
         Shortcut        =   {F7}
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
Dim Oi As Long
Dim Row As String
Dim x As Long
Private Sub Grind_AfterEdit(Row As Integer, Col As Integer, NewValue As String)
EditGrind Col, Row
End Sub

Private Sub Grind_BeforeEdit(Row As Integer, Col As Integer, ByVal Cancel As Boolean)
'If FrmTela(TelaAtual).P.Caption = "1" Then
   Grind.Text = ""
'End If

End Sub


Private Sub Command1_Click()

End Sub

Private Sub Grind_AfterColEdit(ByVal ColIndex As Integer)
EscrevaGrind
End Sub

Private Sub Grind_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode <> vbKeyUp Or KeyCode <> vbKeyDown Then
    If Grind.Col = 0 Then Grind.Col = 1
End If
End Sub

Private Sub MDIForm_Load()
Banco.DatabaseName = NomeBanco
ReDim FrmTela(999) As New Form2
ReDim Run(999) As New Compl
ReDim Menus(999) As New FrmM
MDIForm_Resize
ContTela = 0
TelaAtual = -1
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

Private Sub MenuPropriedade_Click()
Grind.SetFocus
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
Field "Form" + Trim(Str(ContTela)), 1
Banco.DatabaseName = NomeBanco
Banco.RecordSource = "Form" + Trim(Str(ContTela))
Banco.Refresh
With Banco.Recordset
    .Edit
    !Campo2 = A
    .Update
    .MoveNext
    .Edit
    !Campo2 = FrmTela(ContTela).BackColor
    .Update
    .MoveNext
    .Edit
    !Campo2 = FrmTela(ContTela).Caption
    .Update
    .MoveNext
    .Edit
    !Campo2 = FrmTela(ContTela).Height
    .Update
    .MoveNext
    .Edit
    !Campo2 = FrmTela(ContTela).Width
    .Update
    .MoveNext
    .Edit
    !Campo2 = "Não"
    .Update

End With
Banco.Refresh
ContTela = ContTela + 1
Grind.DefColWidth = 1050
Grind.ReBind
End Sub

Private Sub MenuTelas_Click()
Prog.SetFocus

End Sub

Private Sub Prog_DblClick()
If Prog.ListIndex > -1 Then
    FrmTela(Prog.ListIndex).Show vbModeless
End If
End Sub

Private Sub T_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim x As Long
FrmTela(TelaAtual).MousePointer = 2
For x = 1 To T.Buttons.Count
    T.Buttons(x).Value = tbrUnpressed
Next x
Button.Value = tbrPressed
Select Case Button.Key
    Case Is = "B"
        FrmTela(TelaAtual).O.Caption = "1"
    Case Is = "C"
        FrmTela(TelaAtual).O.Caption = "2"
    Case Is = "D"
        FrmTela(TelaAtual).O.Caption = "3"
    Case Is = "E"
        FrmTela(TelaAtual).O.Caption = "4"
    Case Is = "F"
        FrmTela(TelaAtual).O.Caption = "5"
    Case Is = "G"
        FrmTela(TelaAtual).O.Caption = "6"
    Case Is = "H"
        FrmTela(TelaAtual).O.Caption = "7"
    Case Is = "J"
        FrmTela(TelaAtual).O.Caption = "8"
    Case Else
        FrmTela(TelaAtual).MousePointer = 1
        FrmTela(TelaAtual).O.Caption = "0"
    End Select
End Sub



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

Public Sub EscrevaGrind()
Dim Row As Long
On Error Resume Next
If Left(UCase(Banco.RecordSource), 4) = "FORM" And Len(Banco.RecordSource) < 6 Then
    Grind.Col = 1
    Row = Grind.Row + 1
    If Row = 1 Then
        For I = 0 To Prog.ListCount
            If UCase(FrmTela(TelaAtual).Nome.Caption) = UCase(Prog.List(I)) Then
                Prog.List(I) = Grind.Text
                FrmTela(TelaAtual).Nome.Caption = Grind.Text
            End If
        Next I
    ElseIf Row = 2 Then
        FrmTela(TelaAtual).BackColor = Grind.Text
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
            Grind.Columns(1) = "Sim"
            FrmTela(TelaAtual).Max.Caption = Grind.Text
        Else
            FrmTela(TelaAtual).Max.Caption = "Não"
            Grind.Columns(1) = "Não"
        End If
        Grind.Refresh
    End If
ElseIf InStr(1, Banco.RecordSource, "Cmd") <> 0 Then
    Grind.Col = 1
    Row = Grind.Row + 1
    With FrmTela(TelaAtual).Cmd(IndexObj)
        If Row = 1 Then
            .ToolTipText = Grind.Text
        ElseIf Row = 2 Then
            .BackColor = Grind.Text
        ElseIf Row = 3 Then
            .Caption = Grind.Text
        ElseIf Row = 4 Then
            .Height = Grind.Text
        ElseIf Row = 5 Then
            .Width = Grind.Text
        ElseIf Row = 6 Then
            .Top = Grind.Text
        ElseIf Row = 7 Then
            .Left = Grind.Text
        ElseIf Row = 8 Then
            .FontName = Grind.Text
        End If
    End With
ElseIf InStr(1, Banco.RecordSource, "Lbl") <> 0 Then
    Grind.Col = 1
    Row = Grind.Row + 1
    With FrmTela(TelaAtual).Lbl(IndexObj)
        If Row = 1 Then
            .ToolTipText = Grind.Text
        ElseIf Row = 2 Then
            .BackColor = Grind.Text
        ElseIf Row = 3 Then
            .ForeColor = Grind.Text
        ElseIf Row = 4 Then
            .Caption = Grind.Text
        ElseIf Row = 5 Then
            .Height = Grind.Text
        ElseIf Row = 6 Then
            .Width = Grind.Text
        ElseIf Row = 7 Then
            .Top = Grind.Text
        ElseIf Row = 8 Then
            .Left = Grind.Text
        ElseIf Row = 9 Then
            .FontName = Grind.Text
        ElseIf Row = 10 Then
            .FontSize = Grind.Text
        End If
    End With
ElseIf InStr(1, UCase(Banco.RecordSource), "TXT") <> 0 Then
    Grind.Col = 1
    Row = Grind.Row + 1
    With FrmTela(TelaAtual).Txt(IndexObj)
        If Row = 1 Then
            .ToolTipText = Grind.Text
        ElseIf Row = 2 Then
            .BackColor = Grind.Text
        ElseIf Row = 3 Then
            .ForeColor = Grind.Text
        ElseIf Row = 4 Then
            .Text = Grind.Text
        ElseIf Row = 5 Then
            .Height = Grind.Text
        ElseIf Row = 6 Then
            .Width = Grind.Text
        ElseIf Row = 7 Then
            .Top = Grind.Text
        ElseIf Row = 8 Then
            .Left = Grind.Text
        ElseIf Row = 9 Then
            .FontName = Grind.Text
        ElseIf Row = 10 Then
            .FontSize = Grind.Text
        End If
    End With
ElseIf InStr(1, UCase(Banco.RecordSource), "FM") <> 0 Then
Grind.Col = 1
    Row = Grind.Row + 1
    With FrmTela(TelaAtual).Fm(IndexObj)
        If Row = 1 Then
            .ToolTipText = Grind.Text
        ElseIf Row = 2 Then
            .BackColor = Grind.Text
        ElseIf Row = 3 Then
            .Caption = Grind.Text
        ElseIf Row = 4 Then
            .Height = Grind.Text
        ElseIf Row = 5 Then
            .Width = Grind.Text
        ElseIf Row = 6 Then
            .Top = Grind.Text
        ElseIf Row = 7 Then
            .Left = Grind.Text
        ElseIf Row = 8 Then
            .FontName = Grind.Text
        ElseIf Row = 9 Then
            .FontSize = Grind.Text
        End If
    End With
End If
End Sub
