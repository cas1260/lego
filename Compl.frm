VERSION 5.00
Begin VB.Form Compl 
   Caption         =   "Form3"
   ClientHeight    =   5415
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   ScaleHeight     =   5415
   ScaleWidth      =   6000
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Cbo 
      Height          =   315
      Index           =   0
      Left            =   1230
      TabIndex        =   9
      Top             =   1440
      Visible         =   0   'False
      Width           =   1665
   End
   Begin VB.Frame Fm 
      Height          =   525
      Index           =   0
      Left            =   3660
      TabIndex        =   4
      Top             =   2220
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.PictureBox Img 
      Height          =   525
      Index           =   0
      Left            =   2850
      ScaleHeight     =   465
      ScaleWidth      =   1185
      TabIndex        =   3
      Top             =   60
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.TextBox Txt 
      Height          =   285
      Index           =   0
      Left            =   1350
      TabIndex        =   2
      Top             =   60
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.CommandButton Cmd 
      Height          =   525
      Index           =   0
      Left            =   1260
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2520
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.CheckBox Chk 
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.Label A 
      Height          =   375
      Left            =   3270
      TabIndex        =   10
      Top             =   1350
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Lbl 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   870
      TabIndex        =   8
      Top             =   1860
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label Cont 
      Height          =   165
      Left            =   3540
      TabIndex        =   7
      Top             =   4500
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.Label Nome 
      Height          =   345
      Left            =   90
      TabIndex        =   6
      Top             =   4380
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Label P 
      Height          =   405
      Left            =   1140
      TabIndex        =   5
      Top             =   3750
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.Menu M1 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu M1a 
         Caption         =   ""
         Index           =   0
      End
   End
   Begin VB.Menu M2 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu M2a 
         Caption         =   ""
         Index           =   0
      End
   End
   Begin VB.Menu M3 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu M3a 
         Caption         =   ""
         Index           =   0
      End
   End
   Begin VB.Menu M4 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu M4a 
         Caption         =   ""
         Index           =   0
      End
   End
   Begin VB.Menu M5 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu M5a 
         Caption         =   ""
         Index           =   0
      End
   End
   Begin VB.Menu M6 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu M6a 
         Caption         =   ""
         Index           =   0
      End
   End
   Begin VB.Menu M7 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu M7a 
         Caption         =   ""
         Index           =   0
      End
   End
   Begin VB.Menu M8 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu M8a 
         Caption         =   ""
         Index           =   0
      End
   End
End
Attribute VB_Name = "Compl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Click()
If IsNull(Ev_1Vezes(A.Caption)) = False Then
    Dim I As String
    I = Ev_1Vezes(A.Caption)
    WinComandos Ev_1Vezes(A.Caption)
    Ev_1Vezes(A.Caption) = I
End If
End Sub

Private Sub Form_DblClick()
If IsNull(Ev_2Vezes(A.Caption)) = False Then
    WinComandos Ev_2Vezes(A.Caption)
End If
End Sub

Private Sub Form_GotFocus()
If IsNull(Ev_Ganhar(A.Caption)) = False Then
    WinComandos Ev_Ganhar(A.Caption)
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If IsNull(Ev_Escrever(A.Caption)) = False Then
    WinComandos Ev_Escrever(A.Caption)
End If
End Sub

Private Sub Form_Load()
ReDim NomeVar(1000) As String, ValorVar(1000) As String, TipoVar(1000) As String
If Trim(A.Caption) <> "" Then
    If IsNull(Ev_Load(A.Caption)) = False Then
        WinComandos Ev_Load(A.Caption)
    End If
End If
End Sub
Private Sub Form_LostFocus()
If IsNull(Ev_Peder(A.Caption)) = False Then
    WinComandos Ev_Peder(A.Caption)
End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If IsNull(Ev_Fechar(A.Caption)) = False Then
    WinComandos Ev_Fechar(A.Caption)
End If
End Sub

Private Sub Form_Resize()
If IsNull(Ev_Red(A.Caption)) = False Then
    WinComandos Ev_Red(A.Caption)
End If
End Sub
Private Function WinComandos(Com As String) As Boolean
On Error Resume Next
Dim C As Long, Inst As String, X As Long
Dim Inst1 As String, Inst2 As String
Dim Y As Long, TAux As String, K As Long, Fim As Boolean
C = 1
TAux = Com
Fim = True

UltVar = 0
FimProg = True
WinComandos = False
Do While C <= Len(Com)
    X = InStr(1, Com, Chr(10))
    K = X
    If X = 0 Then
        X = InStr(C, Com, ":")
        If X = 0 Then
            Inst = Com
        End If
        If Fim = False Then
            Exit Do
        End If
        Fim = False
    Else
        Y = InStr(X + 1, Com, Chr(10))
        If Y = 0 Then
            Y = Len(Com) - X
        End If
        Inst = Left(Com, X)
    End If
    If Inst <> Chr(10) And Inst <> Chr(8) And Inst <> Chr(11) And Inst <> Chr(13) And Inst <> "" And Inst <> Chr(13) + Chr(10) Then
        Y = 0
        Do While True
            If Y = 0 Then
                X = InStr(1, Inst, " ")
            ElseIf Y = 1 Then
                X = InStr(1, Inst, "(")
            ElseIf Y = 2 Then
                X = InStr(1, Inst, "[")
            ElseIf Y = 3 Then
                X = InStr(1, Inst, "{")
            ElseIf Y = 4 Then
                X = InStr(1, Inst, "=")
            End If
            If X <> 0 Then
                Inst1 = Left(Inst, X)
                Inst2 = Right(Inst, Len(Inst) - X)
                Exit Do
            End If
            Y = Y + 1
            If Y = 5 Then
                Exit Do
            End If
        Loop
        Inst1 = Trim(Elimina(Inst1))
        Inst2 = Trim(Elimina(Inst2))
        'MsgBox Right(Com, Len(Com) - K)
        Com = Right(Com, Len(Com) - K)
        Select Case UCase(Inst1)
            Case "MEM"
                If Mem(Inst1, Inst2) = False Then
                    Exit Function
                End If
            Case "MSG"
                If Msg(Inst1, Inst2) = False Then
                    WinComandos = True
                    Exit Function
                End If
            Case "LOOP"
            Case Else
                If AddVar(Inst1, Inst2) = False Then
                    WinComandos = True
                    Exit Function
                End If
        End Select
    Else
        Inst1 = Left(Com, 5)
        Inst2 = Right(Com, Len(Com) - 5)
        Com = Elimina(Inst1) + Inst2
    End If
Loop
End Function

Private Function BuscaVarErr(Nome As String) As String
Dim X As Long
On Error Resume Next
Nome = Elimina(Nome)
Do While X <> UltVar
    If Trim(UCase(Nome)) = Trim(UCase(NomeVar(X))) Then
        BuscaVarErr = ValorVar(X)
        Exit Function
    End If
    X = X + 1
Loop
BuscaVarErr = "#ERRO#"
End Function

Private Function AddVar(Par As String, Par1 As String) As Boolean
On Error Resume Next
Dim Indice As Long, X As Long, K As Boolean
Dim UI As Long, V1 As String, V2 As String, V3 As String
Dim Conta As Long, Rest As String, Est As String

Indice = BuscaIndece(Par, True)
AddVar = True
If Indice = -1 Then
    AddVar = False
    Exit Function
End If
Par1 = EliminaVar(Par1)
ValorVar(Indice) = BuscaValor(Par1)

End Function
Private Function Elimina(Texto As String)
On Error Resume Next
Dim X As Long, A As Long
Dim T As String
X = 1

Do While X < Len(Texto) + 1
    A = Asc(Mid(Texto, X, 1))
    If A = 8 Or A = 13 Or A = 10 Or A = 9 Then
        T = Mid(Texto, 1, X - 1)
        T = T + Right(Texto, Len(Texto) - X)
        Texto = T
    End If
    X = X + 1
Loop
Elimina = Texto
End Function

Private Function Mem(Par As String, Par1 As String) As Boolean
On Error Resume Next
Dim G As Long, Ty As String
G = InStr(1, UCase(Par1), "COMO")
Mem = True
Par1 = Elimina(Par1)
If G = 0 Then
    FrmErro.Erro.Caption = "Comando Invalido [ " + Par + " " + Par1 + " ]" + Chr(13) + "Esta faltando Paramento"
    FimProg = False
    FrmErro.Show 1
    If FimProg = False Then
        Mem = False
        Exit Function
    End If
End If
Ty = Right(Par1, Len(Par1) - (G + 3))
If Trim(Ty) = "" Then
    FrmErro.Erro.Caption = "Comando Invalido [ " + Par + " " + Par1 + " ]" + Chr(13) + "Tipo não especisicado"
    FimProg = False
    FrmErro.Show 1
    If FimProg = False Then
        Mem = False
        Exit Function
    End If
End If
NomeVar(UltVar) = Trim(Left(Par1, G - 1))
Ty = Trim(Right(Par1, Len(Par1) - (G + 3)))
TipoVar(UltVar) = Trim(Ty)
UltVar = UltVar + 1
End Function

Private Function Msg(Par As String, Par1 As String) As Boolean
On Error Resume Next
Dim T As String
T = EliminaVar(Par1)
'If Left(Par1, 1) = Chr(32) Then
'    T = Mid(2, Par1, Len(Par1) - 2)
'Else
'    T = BuscaVar(Par1)
'    If T = "#ERRO#" Then
'        Msg = False
'        Exit Function
'        Msg = False
'    End If
'End If
MsgBox BuscaValor(T), vbInformation
Msg = True
End Function

Private Function BuscaVar(Nome As String) As String
Dim X As Long
On Error Resume Next
Nome = Elimina(Nome)
Do While X <> UltVar
    If Trim(UCase(Nome)) = Trim(UCase(NomeVar(X))) Then
        BuscaVar = ValorVar(X)
        Exit Function
    End If
    X = X + 1
Loop
BuscaVar = "#ERRO#"
FrmErro.Erro.Caption = "Comando Invalido [ " + Par + " " + Par1 + " ]" + Chr(13) + "Variavel não declara !"
FimProg = False
FrmErro.Show 1
If FimProg = True Then
    BuscaVar = ""
End If

End Function

Private Function AddVar1(Par As String, Par1 As String) As Boolean
On Error Resume Next
Dim Indice As Long, X As Long, K As Boolean
K = False
X = Len(Par)
Par = EliminaVar(Par)
For X = 0 To UltVar
    If UCase(Par) = UCase(NomeVar(X)) Then
        K = True
        Indice = X
        Exit For
    End If
Next X
If K = False Then
    FrmErro.Erro.Caption = "Comando Invalido [ " + Par + " " + Par1 + " ]" + Chr(13) + "Variavel não declara !"
    FimProg = False
    FrmErro.Show 1
    If FimProg = False Then
        Exit Function
    End If
End If
Par1 = EliminaVar(Par1)
Dim UI As Long, V1 As String, V2 As String, V3 As String
Dim Conta As Long
UI = InStr(1, Par1, " ")
If UI = 0 Then
    V1 = Trim(Par1)
Else
    V1 = Trim(Left(Par1, UI))
End If
If Left(V1, 1) = Chr(34) And Right(V1, 1) = Chr(32) Then
    V2 = Mid(V1, 2, Len(V2) - 1)
ElseIf IsNumeric(V1) Then
    V2 = V1
Else
    V2 = BuscaVarErr(V1)
End If
V1 = V2
If V2 = "#ERRO#" Then
    V2 = V1
End If
UI = InStr(1, Par1, "+")
Conta = 1
If UI = 0 Then
    UI = InStr(1, Par1, "/")
    Conta = 2
    If UI = 0 Then
        UI = InStr(1, Par1, "*")
        Conta = 3
        If UI = 0 Then
            UI = InStr(1, Par1, "-")
            Conta = 4
            If UI = 0 Then
                Conta = 5
            End If
        End If
    End If
End If
If Conta <> 5 Then
    V2 = Trim(Right(Par1, Len(Par1) - UI))
    If Left(V2, 1) = Chr(34) And Right(V2, 1) = Chr(32) Then
        V3 = Mid(V1, 2, Len(V2) - 1)
    ElseIf IsNumeric(V2) Then
        V3 = V2
    Else
        V3 = BuscaVarErr(V2)
    End If
    
    If V3 = "#ERRO#" Then
        V3 = V2
    End If
    If Conta = 1 Then
        V2 = CCur(V1) + CCur(V3)
    ElseIf Conta = 2 Then
        V2 = CCur(V1) / CCur(V3)
    ElseIf Conta = 3 Then
        V2 = CCur(V1) * CCur(V3)
    ElseIf Conta = 4 Then
        V2 = CCur(V1) - CCur(V3)
    End If
Else
    V2 = Par1
End If
ValorVar(Indice) = EliminaVar(V2)

End Function
Private Function EliminaVar(Texto As String)
On Error Resume Next
Dim X As Long, A As Long
Dim T As String
X = 1

Do While X < Len(Texto) + 1
    A = Asc(Mid(Texto, X, 1))
    If A = 8 Or A = 13 Or A = 10 Or A = 9 Or A = Asc("=") Then  'Or A = Asc(" ") Then
        T = Mid(Texto, 1, X - 1)
        T = T + Right(Texto, Len(Texto) - X)
        Texto = T
    ElseIf A = Asc("+") Or A = Asc("-") Or A = Asc("/") Or A = Asc("*") Then
        T = Mid(Texto, 1, X - 1)
        T = T + " " + Chr(A) + " " + Right(Texto, Len(Texto) - X)
        Texto = T
        X = X + 1
    End If
    X = X + 1
Loop
If IsNumeric(Trim(Texto)) = False Then
    Texto = Texto
Else
    Texto = Trim(Texto)
End If

EliminaVar = Texto
End Function

Private Function BuscaIndece(Variavel As String, Erro As Boolean)
On Error Resume Next
Dim K As Boolean, X As Long, Indece As Long
K = False
X = Len(Par)
Variavel = Trim(EliminaVar(Variavel))
For X = 0 To UltVar
    If UCase(Variavel) = UCase(NomeVar(X)) Then
        K = True
        Indice = X
        Exit For
    End If
Next X
If Erro = True Then
    If K = False Then
        BuscaIndece = -2
        FrmErro.Erro.Caption = "Comando Invalido [ " + Par + " " + Par1 + " ]" + Chr(13) + "Variavel não declara !"
        FimProg = False
        FrmErro.Show 1
        If FimProg = False Then
            BuscaIndece = -1
            Exit Function
        End If
        BuscaIndece = -2
    End If
End If
BuscaIndece = Indice
End Function

Private Function CriarTela()
On Error Resume Next
Dim Tabela As String
Dim RsTela As Recordset
Dim ContVar() As String, K As Long

Tabela = "Form0"
Set RsTela = BDLego.OpenRecordset(Tabela, dbOpenSnapshot)
If RsTela.RecordCount = 0 Then
    MsgBox "O Arquivo de Configurações esta Com defeito !!!", vbExclamation, App.Title
    RsTela.Close
    Exit Function
End If

Run(0).Tag = RsTela!Campo2
RsTela.MoveNext
Run(0).BackColor = RsTela!Campo2
RsTela.MoveNext
Run(0).Caption = RsTela!Campo2
RsTela.MoveNext
Run(0).Height = RsTela!Campo2
RsTela.MoveNext
Run(0).Width = RsTela!Campo2
RsTela.MoveNext
If UCase(Trim(RsTela!Campo2)) = "SIM" Then
    Run(0).WindowState = 2
Else
    Run(0).WindowState = 0
End If
Dim X As Long, Y As Long
ReDim ContVar(12) As String
X = 0
For X = 0 To 4
    For Y = 1 To 100
        If X = 0 Then
            Tabela = "Form0Cmd" + Trim(Str(Y))
        ElseIf X = 1 Then
            Tabela = "Form0Txt" + Trim(Str(Y))
        ElseIf X = 2 Then
            Tabela = "Form0Lbl" + Trim(Str(Y))
        ElseIf X = 3 Then
            Tabela = "Form0Chk" + Trim(Str(Y))
        ElseIf X = 4 Then
            Tabela = "Form0Cbo" + Trim(Str(Y))
        End If
        If PesqTabela(Tabela) = True Then
            Set RsTela = BDLego.OpenRecordset(Tabela, dbOpenSnapshot)
            If RsTela.RecordCount <> 0 Then
                K = 0
                Do While Not RsTela.EOF
                   ContVar(K) = RsTela!Campo2
                   RsTela.MoveNext
                    K = K + 1
                Loop
                If X = 0 Then
                    RsTela.MoveFirst
                    Load Run(0).Cmd(Run(0).Cmd.Count)
                    With Run(0).Cmd(Run(0).Cmd.Count - 1)
                        .Tag = ContVar(0)
                        .Caption = ContVar(2)
                        .Height = ContVar(3)
                        .Width = ContVar(4)
                        .Top = ContVar(5)
                        .Left = ContVar(6)
                        .FontName = ContVar(7)
                        .Visible = True
                    End With
                ElseIf X = 1 Then
                    Load Run(0).Txt(Run(0).Txt.Count)
                    With Run(0).Txt(Run(0).Txt.Count - 1)
                        .Tag = ContVar(0)
                        .BackColor = ContVar(1)
                        .ForeColor = ContVar(2)
                        .Text = Trim(ContVar(3))
                        .Height = ContVar(4)
                        .Width = ContVar(5)
                        .Top = ContVar(6)
                        .Left = ContVar(7)
                        .FontName = ContVar(8)
                        .FontSize = ContVar(9)
                        .Visible = True
                    End With
                ElseIf X = 2 Then
                    Load Run(0).Lbl(Run(0).Lbl.Count)
                    With Run(0).Lbl(Run(0).Lbl.Count - 1)
                        .Tag = ContVar(0)
                        .BackColor = ContVar(1)
                        .ForeColor = ContVar(2)
                        .Caption = ContVar(3)
                        .Height = ContVar(4)
                        .Width = ContVar(5)
                        .Top = ContVar(6)
                        .Left = ContVar(7)
                        .FontName = ContVar(8)
                        .FontSize = ContVar(9)
                        .Visible = True
                    End With
                ElseIf X = 3 Then
                    Load Run(0).Chk(Run(0).Chk.Count)
                    With Run(0).Chk(Run(0).Chk.Count - 1)
                        .Tag = ContVar(0)
                        .BackColor = ContVar(1)
                        .ForeColor = ContVar(2)
                        .Caption = ContVar(3)
                        .Height = ContVar(4)
                        .Width = ContVar(5)
                        .Top = ContVar(6)
                        .Left = ContVar(7)
                        .FontName = ContVar(8)
                        .FontSize = ContVar(9)
                        .Visible = True
                    End With
                End If
            End If
        End If
    Next Y
Next X
Run(0).Show 1
End Function

Private Function BuscaValor(Par1 As String)
On Error Resume Next
Dim V1 As String, V2 As String, V3 As String

UI = InStr(1, Trim(Par1), " ")
Par1 = Trim(Par1)

If Left(Par1, 1) = Chr(34) And Right(Par1, 1) = Chr(34) Then
    Par1 = Mid(Par1, 2, Len(Par1) - 2)
    BuscaValor = Par1
    Exit Function
End If
If UI = 0 Then
    V1 = Trim(Par1)
Else
    V1 = (Left(Par1, UI))
End If

If Left(V1, 1) = Chr(34) And Right(V1, 1) = Chr(32) Then
    V2 = Mid(V1, 2, Len(V2) - 1)
ElseIf IsNumeric(V1) Then
    V2 = V1
Else
    V2 = BuscaVarErr(V1)
End If

V1 = V2

If V2 = "#ERRO#" Then
    V2 = V1
End If

UI = InStr(1, Par1, "+")

Conta = 1

If UI = 0 Then
    UI = InStr(1, Par1, "/")
    Conta = 2
    If UI = 0 Then
        UI = InStr(1, Par1, "*")
        Conta = 3
        If UI = 0 Then
            UI = InStr(1, Par1, "-")
            Conta = 4
            If UI = 0 Then
                Conta = 5
            End If
        End If
    End If
End If
If Conta <> 5 Then
    V2 = Trim(Right(Par1, Len(Par1) - UI))
    If Left(V2, 1) = Chr(34) And Right(V2, 1) = Chr(34) Then
        V3 = Mid(V2, 2, Len(V2) - 2)
    ElseIf IsNumeric(V2) Then
        V3 = V2
    Else
        V3 = BuscaVarErr(V2)
    End If
    
    If V3 = "#ERRO#" Then
        V3 = V2
    End If
    If Conta = 1 Then
        If IsNumeric(V1) Then
            V2 = CCur(V1) + CCur(V3)
        Else
            V2 = V1 + V3
        End If
    ElseIf Conta = 2 Then
        V2 = CCur(V1) / CCur(V3)
    ElseIf Conta = 3 Then
        V2 = CCur(V1) * CCur(V3)
    ElseIf Conta = 4 Then
        V2 = CCur(V1) - CCur(V3)
    End If
'Else
'    V2 = Par
End If
BuscaValor = V2
End Function
Private Function PesqTabela(Nome As String) As Boolean
Dim X As Long
PesqTabela = False
Set BDLego = OpenDatabase(NomeBanco)
For X = 0 To BDLego.TableDefs.Count - 1
    If UCase(Nome) = UCase(BDLego.TableDefs(X).Name) Then
        PesqTabela = True
        Exit For
    End If
Next X
End Function

Private Function BuscaV(Par1 As String)

End Function

