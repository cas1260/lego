VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "SS32X25.OCX"
Begin VB.Form FrmPropriedade 
   Caption         =   "Propriedade"
   ClientHeight    =   5130
   ClientLeft      =   3495
   ClientTop       =   525
   ClientWidth     =   3240
   Icon            =   "FrmPropriedade.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   3240
   Begin VB.PictureBox Box 
      Height          =   3975
      Left            =   210
      ScaleHeight     =   3915
      ScaleWidth      =   2925
      TabIndex        =   1
      Top             =   1110
      Width           =   2985
      Begin FPSpread.vaSpread Grid 
         Height          =   3645
         Left            =   300
         OleObjectBlob   =   "FrmPropriedade.frx":0E42
         TabIndex        =   2
         Top             =   240
         Width           =   2865
      End
      Begin VB.ListBox L 
         Appearance      =   0  'Flat
         Height          =   615
         ItemData        =   "FrmPropriedade.frx":127E
         Left            =   384
         List            =   "FrmPropriedade.frx":1280
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   1440
         Width           =   1764
      End
      Begin VB.CommandButton Botao 
         Height          =   216
         Left            =   840
         Picture         =   "FrmPropriedade.frx":1282
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   504
         Visible         =   0   'False
         Width           =   300
      End
   End
   Begin VB.ComboBox CboObj 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   780
      Width           =   2520
   End
   Begin MSComDlg.CommonDialog Com 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "FrmPropriedade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Req As Boolean, KeyM As Boolean
Dim Controle As Long
Dim Ob As Object
Dim BancoAux As Database
Dim TextoAtingo As String
Private Sub Command1_Click()
Grid.Col = 2
Grid.Row = 1
Grid.Text = Text1.Text
End Sub

Private Sub Botao_Click()
On Error Resume Next
Grid.Col = 2
TextoAtingo = Grid.Text
L.Top = Botao.Top + Botao.Height
L.Left = Me.Width / 2 - 420 '- L.Width
L.Visible = True
L.ZOrder 0
L.Clear
If Controle = 1 Then
    For Each Ob In FrmTela(TelaAtual)
        If UCase("BancoImg") = UCase(Ob.Name) Then
            If Ob.Tag <> "" Then
                L.AddItem Ob.Tag
            End If
        End If
    Next
ElseIf Controle = 2 Then
    AddLista
ElseIf Controle = 3 Then
    L.Clear
    L.AddItem "Sim"
    L.AddItem "Não"
End If
Grid.Col = 2
L.Text = Grid.Text
L.SetFocus

End Sub

Private Sub CboObj_Click()
If CboObj.ListIndex > -1 Then
    For Each Ob In FrmTela(TelaAtual)
        If UCase(Ob.Tag) = UCase(CboObj.Text) Then
            Set NovoObj = Ob
            Selecione1
            Redimesiona
            Exit Sub
        End If
    Next
    If UCase(CboObj.Text) = UCase(FrmTela(TelaAtual).Tag) Then
        Set NovoObj = FrmTela(TelaAtual)
        Selecione1
        Redimesiona
    End If
    If Trim(Nome) <> "" Then
        FrmPropriedade.CboObj.Text = Nome
    End If
End If
End Sub

Private Sub Form_Load()
Req = True
KeyM = True
Me.Height = (FrmPrincipal.ScaleHeight / 2) + 100
Me.Width = 3360
FrmPropriedade.Left = FrmPrincipal.ScaleWidth - FrmPropriedade.Width
FrmPropriedade.Top = (FrmPrincipal.ScaleHeight - FrmPropriedade.Height)
Req = False
If CRun.OpenExe = 0 Then FrmPrincipal.Visible = False
End Sub

Private Sub Form_Resize()
If Me.WindowState <> 1 And Req = False Then
    CboObj.Top = 0
    CboObj.Left = 0
    CboObj.Width = Me.ScaleWidth
    Box.Left = 0
    Box.Top = CboObj.Height + 10
    Box.Height = Me.Height - 380 - CboObj.Height
    Box.Width = Me.ScaleWidth
    Grid.Left = 0
    Grid.Top = 0
    Grid.Height = Box.Height
    Grid.Width = Box.Width
    Grid.ColWidth(1) = Grid.Width / 281
    Grid.ColWidth(2) = Grid.Width / 245
    If CRun.OpenExe = 0 Then FrmPrincipal.Visible = False
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Cancel = 1
Me.Visible = False
End Sub

Private Sub Grid_Change(ByVal Col As Long, ByVal Row As Long)
On Error Resume Next
Dim TE As String

Grid.Row = Row
Grid.Col = 1
TE = Grid.Text
Grid.Col = 2
Select Case UCase(TE)
    Case "LEGENDA"
        NovoObj.Caption = Grid.Text
    Case "COR DE FUNDO", "COR FUNDO"
        NovoObj.BackColor = Grid.Text
    Case "COR DA LETRA", "COR LETRA"
        NovoObj.ForeColor = Grid.Text
    Case "LOCAL"
        Bancos(NovoObj.Index).Local = Grid.Text
    Case "NOME"
        Dim Xp As Long, Passa As Boolean, Verrifica As Boolean
        Dim X As Long
        Passa = True: Verrifica = False
        
        If UCase(Grid.Text) = UCase(NovoObj.Tag) Then
            Verrifica = True
        End If
        
        If Trim(Grid.Text) = "" Then
            MsgBox "Nome Invalido !!!", vbCritical, App.Title
            Grid.Text = NovoObj.Tag
            Exit Sub
        End If
        
        For Xp = 1 To FrmCodigo.Eventos.Nodes.Count
'            If UCase(FrmCodigo.Eventos.Nodes(Xp).Text) = UCase(Grid.Text) Then
            If UCase(FrmCodigo.Eventos.Nodes(Xp).Key) = UCase(FrmTela(TelaAtual).Tag + "." + Grid.Text) Then
                Passa = False
                Exit For
            End If
        Next Xp
        If Passa = False Then
            If Verrifica = False Then
                MsgBox "Nome ja existente ! ! !", vbCritical, App.Title
                Grid.Text = NovoObj.Tag
                Exit Sub
            End If
        End If
        
        If NovoObj.Name <> "Form2" Then
            For Xp = 1 To FrmCodigo.Eventos.Nodes.Count
                If UCase(FrmCodigo.Eventos.Nodes(Xp).Text) = UCase(NovoObj.Tag) Then
                    FrmCodigo.Eventos.Nodes(Xp).Text = Grid.Text
                    Nome = FrmCodigo.Eventos.Nodes(Xp).Key
                    FrmCodigo.Eventos.Nodes(Xp).Key = FrmTela(TelaAtual).Tag + "." + Grid.Text
                    For X = 1 To FrmCodigo.Eventos.Nodes.Count
                        FrmCodigo.Eventos.Nodes(X).Key = Replace(UCase(FrmCodigo.Eventos.Nodes(X).Key), UCase(Nome), FrmCodigo.Eventos.Nodes(Xp).Key)
                    Next X
                    Passa = False
                    Exit For
                End If
            Next Xp
        End If
        If NovoObj.Name = "Form2" Then
            Passa = True
            For Xp = 1 To FrmEx.Prog.Nodes.Count
                If UCase(FrmEx.Prog.Nodes(Xp).Text) = UCase(Grid.Text) Then
                    Passa = False
                    Exit For
                End If
            Next Xp
            If Passa = False Then
                If Verrifica = False Then
                    MsgBox "Nome ja existente ! ! !", vbCritical, App.Title
                    Grid.Text = NovoObj.Tag
                    Exit Sub
                End If
            End If
            
            For Xp = 1 To FrmEx.Prog.Nodes.Count
                If UCase(FrmEx.Prog.Nodes(Xp).Text) = UCase(NovoObj.Tag) Then
                    FrmEx.Prog.Nodes(Xp).Text = Grid.Text
                    FrmEx.Prog.Nodes(Xp).Key = Grid.Text
                    Passa = False
                    Exit For
                End If
            Next Xp
            For Xp = 1 To FrmCodigo.Eventos.Nodes.Count
                If UCase(FrmCodigo.Eventos.Nodes(Xp).Text) = UCase(NovoObj.Tag) Then
                    FrmCodigo.Eventos.Nodes(Xp).Text = Grid.Text
                    Nome = UCase(FrmCodigo.Eventos.Nodes(Xp).Key) + "."
                    FrmCodigo.Eventos.Nodes(Xp).Key = UCase(Grid.Text)
                    For X = 1 To FrmCodigo.Eventos.Nodes.Count
                        'If UCase(FrmCodigo.Eventos.Nodes(X).Key + ".") = Left(UCase(Nome), Len(FrmCodigo.Eventos.Nodes(X).Key + ".")) Then
                        If Nome = UCase(Left(FrmCodigo.Eventos.Nodes(X).Key, Len(Nome))) Then
                            FrmCodigo.Eventos.Nodes(X).Key = Replace(UCase(FrmCodigo.Eventos.Nodes(X).Key), UCase(Nome), FrmCodigo.Eventos.Nodes(Xp).Key)
                        End If
                    Next X
                    Passa = False
                    Exit For
                End If
            Next Xp

        End If
        NovoObj.Tag = Grid.Text
    Case "TAMX"
        NovoObj.Height = Grid.Text
    Case "TAMY"
        NovoObj.Width = Grid.Text
    Case "POXX"
        NovoObj.Top = Grid.Text
    Case "POXY"
        NovoObj.Left = Grid.Text
    Case "FONTE"
        NovoObj.FontName = Grid.Text
    Case "TAMANHO"
        NovoObj.FontSize = Grid.Text
    Case "TEXTO"
        NovoObj.Text = Grid.Text
    Case "IMAGEM"
        If Dir(Grid.Text) = "" Then
            MsgBox "Imagem Invalida ! ! !", vbCritical, App.Title
            Exit Sub
        End If
        NovoObj.ToolTipText = Grid.Text
    Case "BORDA"
        If UCase(Left(Grid.Text, 1)) = "S" Then
            Grid.Text = "Sim"
            NovoObj.BorderStyle = 1
        Else
            NovoObj.BorderStyle = 0
            Grid.Text = "Não"
        End If
    Case "COMPRIMIR"
        If UCase(Left(Grid.Text, 1)) = "S" Then
            Grid.Text = "Sim"
            NovoObj.Stretch = 1
        Else
            NovoObj.Stretch = 0
            Grid.Text = "Não"
        End If
     
    Case "ORDEM"
        If IsNumeric(Grid.Text) = True Then
            NovoObj.TabIndex = Grid.Text
        Else
            Grid.Text = NovoObj.TabIndex
        End If
    Case "3D"
        If UCase(Left(Grid.Text, 1)) = "S" Then
            Grid.Text = "Sim"
            NovoObj.Appearance = 1
        Else
            NovoObj.Appearance = 0
            Grid.Text = "Não"
        End If
    Case "MASCARA"
        NovoObj.PasswordChar = Left(Grid.Text, 1)
        Grid.Text = Left(Grid.Text, 1)
    Case "BOTAO MAX"
        If UCase(Left(Grid.Text, 1)) = "S" Then
            Grid.Text = "Sim"
            Max(NovoObj.Cont.Caption) = True
        Else
            Max(NovoObj.Cont.Caption) = False
            Grid.Text = "Não"
        End If
    Case "BOTAO MIN"
        If UCase(Left(Grid.Text, 1)) = "S" Then
            Grid.Text = "Sim"
            Min(NovoObj.Cont.Caption) = True
        Else
            Min(NovoObj.Cont.Caption) = False
            Grid.Text = "Não"
        End If
    Case "BOTAO FECHAR"
        If UCase(Left(Grid.Text, 1)) = "S" Then
            Grid.Text = "Sim"
            Fecha(NovoObj.Cont.Caption) = True
        Else
            Fecha(NovoObj.Cont.Caption) = False
            Grid.Text = "Não"
        End If
    Case "TABELA"
        Index = NovoObj.Index
        TabRec(Index).Tabela = Grid.Text
    Case "CODIÇÕES"
        Index = NovoObj.Index
        TabRec(Index).Codicao = Grid.Text
    Case "ORDEM"
        Index = NovoObj.Index
        TabRec(Index).Ordem = Grid.Text
    Case "BANCO"
        Index = NovoObj.Index
        TabRec(Index).BancoDb = Grid.Text
End Select
End Sub

Private Sub Grid_Click(ByVal Col As Long, ByVal Row As Long)
On Error Resume Next
Controle = 0
Botao.Visible = False
L.Visible = False
Grid.Col = 1
Grid.Row = Row
'MsgBox Col & " " & Row
If InStr(1, UCase(Grid.Text), "BANCO") <> 0 Then
    MostraBotao 1
ElseIf InStr(1, UCase(Grid.Text), "TABELA") <> 0 Then
    MostraBotao 2
ElseIf VerrificaBotao(Grid.Text) = True Then

ElseIf InStr(1, UCase(Grid.Text), "COR") <> 0 Then
    Com.Color = 0
    Com.ShowColor
    If Com.Color <> 0 Then
        If UCase(Grid.Text) = "COR DE FUNDO" Or UCase(Grid.Text) = "COR FUNDO" Then
            If NovoObj.Name = "Form2" Then
                FrmTela(TelaAtual).Xmenu.BackColor = Com.Color
            End If
            NovoObj.BackColor = Com.Color
        ElseIf UCase(Grid.Text) = "COR DA LETRA" Or UCase(Grid.Text) = "COR LETRA" Then
            NovoObj.ForeColor = Com.Color
        End If
        Grid.Col = 2
        Grid.Text = Com.Color
    End If
ElseIf InStr(1, UCase(Grid.Text), "LOCAL") <> 0 Then
    Com.FileName = ""
    Com.Filter = "Arquivo de Base de dados |*.Mdb;*.mdb |Todos os Arquivo |*.*"
    Com.DialogTitle = "Lego - Banco de dados"
    Com.ShowOpen
    If Com.FileName = "" Then
        Exit Sub
    End If
    If Dir(Com.FileName) = "" Then
        MsgBox "Banco de dados não encontrada ! ! !", vbCritical, App.Title
        Exit Sub
    End If
    Bancos(NovoObj.Index).Local = Com.FileName
    Grid.Col = 2
    Grid.Text = Com.FileName
ElseIf InStr(1, UCase(Grid.Text), "IMAGEM") <> 0 Then
    Com.FileName = ""
    Com.Filter = "Todos os Arquivo de Figuras |*.Bmp;*.Jpeg;*.jpg;*.Gif |Todos os Arquivo |*.*"
    Com.DialogTitle = "Lego - Imagem"
    Com.ShowOpen
    If Com.FileName = "" Then
        Exit Sub
    End If
    If Dir(Com.FileName) = "" Then
        MsgBox "Imagem não encontrada ! ! !", vbCritical, App.Title
        Exit Sub
    End If
    NovoObj.Picture = LoadPicture(Com.FileName)
    Grid.Col = 2
    Grid.Text = Com.FileName
'    NovoObj.ToolTipText = Com.FileName
ElseIf InStr(1, UCase(Grid.Text), "FONTE") <> 0 Then
  Com.flags = cdlCFBoth + cdlCFEffects
  Com.ShowFont
  With NovoObj
    .FontName = Com.FontName
    .FontSize = Com.FontSize
    .FontBold = Com.FontBold
    .FontItalic = Com.FontItalic
    .FontStrikethru = Com.FontStrikethru
    .FontUnderline = Com.FontUnderline
  End With
End If
End Sub

Private Sub Grid_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
If Col = 1 Then
'    Grid.ac
End If
End Sub

Private Sub Grid_KeyPress(KeyAscii As Integer)
'MsgBox "dd"
'If Grid.Col = 2 Then
'    KeyAscii = 0
'End If
End Sub

Private Sub St_Click(PreviousTab As Integer)
If St.Tab = 0 Then
    Grid.Visible = True
    Eventos.Visible = False
ElseIf St.Tab = 1 Then
    Grid.Visible = False
    Eventos.Visible = True
End If
End Sub

Private Sub Grid_KeyUp(KeyCode As Integer, Shift As Integer)
Dim Texto As String
Grid.Col = Grid.Col
Grid.Row = Grid.Row
Texto = Grid.Text
    If InStr(1, UCase(Texto), "BANCO") <> 0 Then
        MostraBotao 1
    ElseIf InStr(1, UCase(Texto), "TABELA") <> 0 Then
        MostraBotao 2
    ElseIf VerrificaBotao(Texto) = True Then
    End If
' End If
End Sub

Private Sub L_Click()
If L.ListIndex <> -1 Then
    Grid.Col = 2
    Grid.Text = L.Text
End If
End Sub

Private Sub L_DblClick()
L_KeyDown 13, 0
End Sub

Private Sub L_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    L_Click
    L.Visible = False
    Botao.Visible = False
    Grid_Change Grid.Col, Grid.Row
End If
If KeyCode = 27 Then
    Grid.Col = 2
    Grid.Text = TextoAtingo
    Grid.SetFocus
    Exit Sub
End If
L_Click
End Sub

Private Sub L_LostFocus()
L.Visible = False
Botao.Visible = False
End Sub

Private Sub AddLista()
On Error Resume Next
Dim Local1 As String
Dim Ban As String
Dim Index As Long
Dim X As Long

Index = 0
Ban = TabRec(NovoObj.Index).BancoDb
For Each Ob In FrmTela(TelaAtual)
    If UCase(Ban) = UCase(Ob.Tag) Then
        Index = Ob.Index
        Exit For
    End If
Next
L.Clear
If Index = 0 Then Exit Sub

Local1 = Bancos(Index).Local

If Dir(Local1) = "" Or Trim(Local1) = "" Then
    MsgBox "Caminho invalido, ou inexistente", vbInformation, App.Title
    Exit Sub
End If
Set BancoAux = OpenDatabase(Local1)
'If BancoAux = Nothing Then Exit Sub
For X = 0 To BancoAux.TableDefs.Count
    If Left(UCase(BancoAux.TableDefs(X).Name), 4) <> "MSYS" And Left(UCase(BancoAux.TableDefs(X).Name), 4) <> "RTBL" Then
        L.AddItem BancoAux.TableDefs(X).Name
    End If
Next X
End Sub

Private Function MostraBotao(Controles As Long)
Botao.Visible = True
Botao.Left = (Me.Width) - Botao.Width - 350
Botao.Top = (Grid.RowHeight(Grid.Row) * 18) * Grid.Row - 229.5
Botao.ZOrder 0
Controle = Controles
End Function

Private Function VerrificaBotao(Texto As String) As Boolean
On Error Resume Next
VerrificaBotao = False
Texto = UCase(Texto)
If Texto = "BORDA" Or Texto = "3D" Or Texto = "BOTAO MAX" Or Texto = "BOTAO MIN" Or Texto = "BOTAO FECHAR" Or Texto = "COMPRIMIR" Then
    Controle = 3
    VerrificaBotao = True
    MostraBotao Controle
End If
End Function
