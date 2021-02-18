Attribute VB_Name = "Geral"

Private Declare Function SendMessageLong Lib "user32" Alias _
        "SendMessageA" (ByVal HWnd As Long, ByVal wMsg As Long, _
        ByVal wParam As Long, lParam As Long) As Long

Private Const EM_LINEFROMCHAR = &HC9
Private Const EM_LINEINDEX = &HBB
Private Const EM_LINELENGTH = &HC1
    
Public Sub Redimesiona()
On Error Resume Next
Dim X As Long, O As Object

Set O = NovoObj
FrmTela(TelaAtual).Im(0).Top = O.Top - FrmTela(TelaAtual).Im(0).Height
FrmTela(TelaAtual).Im(0).Left = O.Left - 100
FrmTela(TelaAtual).Im(1).Left = O.Left + ((O.Width / 2) - FrmTela(TelaAtual).Im(1).Width / 2)
FrmTela(TelaAtual).Im(1).Top = FrmTela(TelaAtual).Im(0).Top
FrmTela(TelaAtual).Im(2).Left = O.Left + O.Width
FrmTela(TelaAtual).Im(2).Top = FrmTela(TelaAtual).Im(1).Top
FrmTela(TelaAtual).Im(6).Left = FrmTela(TelaAtual).Im(0).Left
FrmTela(TelaAtual).Im(6).Top = (O.Top - FrmTela(TelaAtual).Im(6).Height) + (O.Height / 2)

FrmTela(TelaAtual).Im(3).Top = O.Top + O.Height '+ FrmTela(TelaAtual).Im(4).Height
FrmTela(TelaAtual).Im(3).Left = O.Left - 100
FrmTela(TelaAtual).Im(4).Left = FrmTela(TelaAtual).Im(1).Left
FrmTela(TelaAtual).Im(4).Top = FrmTela(TelaAtual).Im(3).Top
FrmTela(TelaAtual).Im(5).Left = O.Left + O.Width
FrmTela(TelaAtual).Im(5).Top = FrmTela(TelaAtual).Im(4).Top
FrmTela(TelaAtual).Im(7).Left = FrmTela(TelaAtual).Im(5).Left
FrmTela(TelaAtual).Im(7).Top = (O.Top - FrmTela(TelaAtual).Im(7).Height) + (O.Height / 2)
If O.Name <> "Form2" Then
    For X = 0 To 7
        FrmTela(TelaAtual).Im(X).Visible = True
        FrmTela(TelaAtual).Im(X).Refresh
    Next X
End If
'FrmTela(TelaAtual).Refresh
'L.Visible = False
'Botao.Visible = False
End Sub


Public Sub Selecione1()
On Error Resume Next
Dim Na() As String, Ny() As String, Ix As Long
Dim Disp As Long, O As Object, TotalGrid As Byte, X As Byte
Dim index As Long

Set O = NovoObj

ReDim Na(15) As String, Ny(15) As String
TotalGrid = 9

With FrmPrincipal

    X = 1
    .Grid.Clear
    .Grid.TextArray(1) = "Nome"
    .Grid.TextArray(2) = "Valores"
    If O.Name = "Txt" Then
        .Grid.Rows = 14
    ElseIf O.Name = "Lst" Then
        .Grid.Rows = 13
    ElseIf O.Name = "Img" Then
       .Grid.Rows = 13
    ElseIf O.Name = "Form2" Then
       .Grid.Rows = 15
    ElseIf O.Name = "BancoImg" Then
        GoTo BancoDeDados
    ElseIf O.Name = "ImgRecord" Then
        GoTo Tabela
    ElseIf O.Name = "Tm" Then
        GoTo Tm1
    Else
        .Grid.Rows = 13
        If O.Name = "Cmd" Then
            .Grid.Rows = .Grid.Rows + 1
        End If
    End If
    .Grid.TextMatrix(X, 1) = "Nome":           .Grid.TextMatrix(X, 2) = NovoObj.Tag
    X = X + 1
    .Grid.TextMatrix(X, 1) = "Borda":          .Grid.TextMatrix(X, 2) = IIf(NovoObj.BorderStyle = 0, "Não", "Sim")
    X = X + 1
    .Grid.TextMatrix(X, 1) = "Objeto 3D":      .Grid.TextMatrix(X, 2) = IIf(NovoObj.Appearance = 0, "Não", "Sim")
    X = X + 1
    .Grid.TextMatrix(X, 1) = "Cor Fundo":      .Grid.TextMatrix(X, 2) = NovoObj.BackColor
    X = X + 1
    .Grid.TextMatrix(X, 1) = "Cor da Letra":   .Grid.TextMatrix(X, 2) = NovoObj.ForeColor
    X = X + 1
    If O.Name = "Txt" Then
        .Grid.TextMatrix(X, 1) = "Texto":      .Grid.TextMatrix(X, 2) = NovoObj.Text
        X = X + 1
        .Grid.TextMatrix(X, 1) = "Mascara":    .Grid.TextMatrix(X, 2) = NovoObj.Text
        X = X + 1
    ElseIf O.Name = "Lst" Then
        .Grid.TextMatrix(X, 1) = "Texto":      .Grid.TextMatrix(X, 2) = NovoObj.Text
        X = X + 1
    ElseIf O.Name = "Img" Then
        .Grid.TextMatrix(X, 1) = "Imagem":     .Grid.TextMatrix(X, 2) = NovoObj.ToolTipText
        X = X + 1
    '    .Grid.Col = 1: .Grid.Row = X:  .Grid.Text = "Borda":      .Grid.Col = 2:  .Grid.Text = IIf(NovoObj.BorderStyle = 0, "Não", "Sim")
    '    X = X + 1
        .Grid.TextMatrix(X, 1) = "Comprimir":  .Grid.TextMatrix(X, 2) = IIf(NovoObj.Stretch = 0, "Não", "Sim")
        X = X + 1
    Else
        .Grid.TextMatrix(X, 1) = "Legenda":   .Grid.TextMatrix(X, 2) = NovoObj.Caption
        X = X + 1
        If O.Name = "Cmd" Then
            .Grid.TextMatrix(X, 1) = "Imagem": .Grid.TextMatrix(X, 2) = NovoObj.ToolTipText
            X = X + 1
        End If
    End If
    If O.Name = "Form2" Then
        .Grid.TextMatrix(X, 1) = "Botao Max":  .Grid.TextMatrix(X, 2) = IIf(Max(NovoObj.Cont.Caption) = True, "Sim", "Não")
        X = X + 1
        .Grid.TextMatrix(X, 1) = "Botao Min":  .Grid.TextMatrix(X, 2) = IIf(Min(NovoObj.Cont.Caption) = True, "Sim", "Não")
        X = X + 1
        .Grid.TextMatrix(X, 1) = "Botao Fechar": .Grid.TextMatrix(X, 2) = IIf(Fecha(NovoObj.Cont.Caption) = True, "Sim", "Não")
        X = X + 1
    End If
    
    If O.Name <> "Form2" And O.Name <> "Img" Then
        .Grid.TextMatrix(X, 1) = "Ordem":    .Grid.TextMatrix(X, 2) = NovoObj.TabIndex
        X = X + 1
    End If
    .Grid.TextMatrix(X, 1) = "TamX":  .Grid.TextMatrix(X, 2) = NovoObj.Height
    X = X + 1
    .Grid.TextMatrix(X, 1) = "TamY": .Grid.TextMatrix(X, 2) = NovoObj.Width
    X = X + 1
    .Grid.TextMatrix(X, 1) = "PoxX":  .Grid.TextMatrix(X, 2) = NovoObj.Top
    X = X + 1
    .Grid.TextMatrix(X, 1) = "PoxY":   .Grid.TextMatrix(X, 2) = NovoObj.Left
    X = X + 1
    .Grid.TextMatrix(X, 1) = "Fonte":    .Grid.TextMatrix(X, 2) = NovoObj.FontName
    X = X + 1
    .Grid.TextMatrix(X, 1) = "Tamanho":  .Grid.TextMatrix(X, 2) = NovoObj.FontSize
    '.Grid.MaxRows = X
End With
For X = 0 To 7
    FrmTela(TelaAtual).Im(X).ZOrder vbSendToBack
    FrmTela(TelaAtual).Im(X).ZOrder vbBringToFront
Next X

Exit Sub

BancoDeDados:
With FrmPrincipal.Grid
    X = 1
    .Rows = 2
    index = NovoObj.index
    .Col = 1: .Row = X: .Text = "Nome": .Col = 2:  .Text = NovoObj.Tag
    X = X + 1
    .Col = 1: .Row = X: .Text = "Local":     .Col = 2:  .Text = Bancos(index).Local
    Bancos(index).Nome = NovoObj.Tag
End With
Exit Sub
Tabela:
With FrmPrincipal.Grid
    X = 1
    .Rows = 4
    index = NovoObj.index
    .Col = 1: .Row = X: .Text = "Nome": .Col = 2:  .Text = NovoObj.Tag
    X = X + 1
    .Col = 1: .Row = X: .Text = "Tabela":     .Col = 2:  .Text = TabRec(index).Tabela
    X = X + 1
    .Col = 1: .Row = X: .Text = "Codições":     .Col = 2:  .Text = TabRec(index).Codicao
    X = X + 1
    .Col = 1: .Row = X: .Text = "Ordem":     .Col = 2:  .Text = TabRec(index).Ordem
    X = X + 1
    .Col = 1: .Row = X: .Text = "Banco": .Col = 2: .Text = TabRec(index).BancoDb
    TabRec(index).Nome = NovoObj.Tag
End With

Exit Sub

Tm1:
With FrmPrincipal
    .Grid.Rows = 3
    X = 1
   .Grid.TextMatrix(X, 1) = "Nome":           .Grid.TextMatrix(X, 2) = NovoObj.Tag
    X = 2
    .Grid.TextMatrix(X, 1) = "Tempo":          .Grid.TextMatrix(X, 2) = NovoObj.DataMember
End With
End Sub
Public Sub SalvarEventos1(Arq As String)
On Error Resume Next
If Dir(Arq) <> "" Then
    Kill Arq
End If
Dim Banco As Database
Dim NewT As TableDef
Dim Rs As Recordset
Dim X As Long

Set Banco = CreateDatabase(Arq, dbLangGeneral, dbEncrypt)
Set NewT = Banco.CreateTableDef("Lego")
NewT.Fields.Append NewT.CreateField("Key", dbText, 50)
NewT.Fields(0).AllowZeroLength = True
NewT.Fields.Append NewT.CreateField("Codigo", dbMemo)
NewT.Fields(1).AllowZeroLength = True
NewT.Fields.Append NewT.CreateField("Texto", dbText, 30)
NewT.Fields(2).AllowZeroLength = True
NewT.Fields.Append NewT.CreateField("Root", dbText, 30)
NewT.Fields(3).AllowZeroLength = True

Banco.TableDefs.Append NewT


Set Rs = Banco.OpenRecordset("Lego", dbOpenDynaset)
For X = 1 To FrmCodigo.Eventos.Nodes.Count
    Rs.AddNew
    Rs!Key = FrmCodigo.Eventos.Nodes(X).Key
    Rs!Codigo = FrmCodigo.Eventos.Nodes(X).Tag
    Rs!Texto = FrmCodigo.Eventos.Nodes(X).Text
    Rs!Root = FrmCodigo.Eventos.Nodes(X).Root
    Rs.Update
Next X
Banco.Close
End Sub
Public Sub BuscaEventosArq(Arq As String)
On Error Resume Next

Dim Rs As Recordset
Dim X As Long
Dim Pesq As String

Pesq = "SELECT * From Lego Where Left(key,5) = 'PROD.'"
Set Rs = Banco.OpenRecordset(Pesq, dbOpenDynaset)
If Rs.EOF = False Then
    Do While Not Rs.EOF
        FrmCodigo.Eventos.Nodes.Add "PROD", tvwChild, Rs!Key, Rs!Texto, 6
        Rs.MoveNext
    Loop
End If
For X = 1 To FrmCodigo.Eventos.Nodes.Count
    Pesq = "Select * From Lego Where Key = '" & FrmCodigo.Eventos.Nodes(X).Key & "'"
    Set Rs = Banco.OpenRecordset(Pesq, dbOpenSnapshot)
    If Rs.EOF = False Then
        FrmCodigo.Eventos.Nodes(X).Key = Rs!Key
        FrmCodigo.Eventos.Nodes(X).Tag = Rs!Codigo
        FrmCodigo.Eventos.Nodes(X).Text = Rs!Texto
        FrmCodigo.Eventos.Nodes(X).Root = Rs!Root
    End If
Next X

Banco.Close
End Sub

Public Sub CriarEstrutura(Tabela As String)
On Error Resume Next

Dim NewT As TableDef
Dim Rs As Recordset
Dim X As Long
Dim Nome() As String, Tipo()   As String, Tamanho() As String


ReDim Nome(21) As String, Tipo(21) As String, Tamanho(21) As String
    
Nome(0) = "Nome": Tipo(0) = dbText: Tamanho(0) = 50
Nome(1) = "Borda": Tipo(1) = dbText: Tamanho(1) = 1
Nome(2) = "3D": Tipo(2) = dbText: Tamanho(2) = 1
Nome(3) = "Cor Fundo": Tipo(3) = dbText: Tamanho(3) = 15
Nome(4) = "Cor da Letra": Tipo(4) = dbText: Tamanho(4) = 15
Nome(5) = "Texto": Tipo(5) = dbText: Tamanho(5) = 60
Nome(6) = "Mascara": Tipo(6) = dbText: Tamanho(6) = 1
Nome(7) = "Imagem": Tipo(7) = dbMemo: Tamanho(7) = ""
Nome(8) = "Comprimir": Tipo(8) = dbText: Tamanho(8) = 3
Nome(9) = "Legenda": Tipo(9) = dbText: Tamanho(9) = 60
Nome(10) = "BotaoMax": Tipo(10) = dbText: Tamanho(10) = 3
Nome(11) = "BotaoMin": Tipo(11) = dbText: Tamanho(11) = 3
Nome(12) = "BotaoFechar": Tipo(12) = dbText: Tamanho(12) = 3
Nome(13) = "Ordem": Tipo(13) = dbText: Tamanho(13) = 3
Nome(14) = "TamX": Tipo(14) = dbText: Tamanho(14) = 8
Nome(15) = "TamY": Tipo(15) = dbText: Tamanho(15) = 8
Nome(16) = "PoxX": Tipo(16) = dbText: Tamanho(16) = 8
Nome(17) = "PoxY": Tipo(17) = dbText: Tamanho(17) = 8
Nome(18) = "Fonte": Tipo(18) = dbText: Tamanho(18) = 30
Nome(19) = "Tamanho": Tipo(19) = dbText: Tamanho(19) = 3
Nome(20) = "Tipo": Tipo(20) = dbText: Tamanho(20) = 50
Nome(21) = "Estilo": Tipo(21) = dbText: Tamanho(21) = 50

Set NewT = Banco.CreateTableDef(Tabela)
For X = 0 To 21
    If Trim(Tamanho(X)) = "" Then
       NewT.Fields.Append NewT.CreateField(Nome(X), Tipo(X))
    Else
        NewT.Fields.Append NewT.CreateField(Nome(X), Tipo(X), Tamanho(X))
    End If
    NewT.Fields(X).AllowZeroLength = True
Next X
Banco.TableDefs.Append NewT

Set NewT = Banco.CreateTableDef("Menu-" + Tabela)
NewT.Fields.Append NewT.CreateField("Legenda", dbText, 50)
NewT.Fields.Append NewT.CreateField("Ident", dbLong, 3)
NewT.Fields.Append NewT.CreateField("Chave", dbText, 30)
NewT.Fields.Append NewT.CreateField("Root", dbText, 30)

NewT.Fields(0).AllowZeroLength = True
NewT.Fields(1).AllowZeroLength = True
NewT.Fields(2).AllowZeroLength = True
NewT.Fields(3).AllowZeroLength = True

Banco.TableDefs.Append NewT

End Sub

Public Sub CriarBanco(Arq As String)


Set Banco = CreateDatabase(Arq, dbLangGeneral, dbEncrypt)

Set NewT = Banco.CreateTableDef("Lego")
NewT.Fields.Append NewT.CreateField("Key", dbText, 50)
NewT.Fields(0).AllowZeroLength = True
NewT.Fields.Append NewT.CreateField("Codigo", dbMemo)
NewT.Fields(1).AllowZeroLength = True
NewT.Fields.Append NewT.CreateField("Texto", dbText, 30)
NewT.Fields(2).AllowZeroLength = True
NewT.Fields.Append NewT.CreateField("Root", dbText, 30)
NewT.Fields(3).AllowZeroLength = True
Banco.TableDefs.Append NewT

Set NewT = Banco.CreateTableDef("Config")
NewT.Fields.Append NewT.CreateField("Iniciar", dbText, 50)
NewT.Fields(0).AllowZeroLength = True
NewT.Fields.Append NewT.CreateField("Plataforma", dbText, 15)
NewT.Fields(1).AllowZeroLength = True
NewT.Fields.Append NewT.CreateField("EXE", dbText, 30)
NewT.Fields(2).AllowZeroLength = True
NewT.Fields.Append NewT.CreateField("Titulo", dbText, 50)
NewT.Fields(3).AllowZeroLength = True
NewT.Fields.Append NewT.CreateField("Autor", dbText, 50)
NewT.Fields(4).AllowZeroLength = True
NewT.Fields.Append NewT.CreateField("Senha", dbText, 50)
NewT.Fields(5).AllowZeroLength = True
NewT.Fields.Append NewT.CreateField("Banco", dbText, 50)
NewT.Fields(6).AllowZeroLength = True
NewT.Fields.Append NewT.CreateField("Esc", dbLong, 1)
NewT.Fields(7).AllowZeroLength = True
NewT.Fields.Append NewT.CreateField("Enter", dbLong, 1)
NewT.Fields(8).AllowZeroLength = True

Banco.TableDefs.Append NewT
Set Banco = OpenDatabase(Arq)
End Sub

Public Sub SalvarEventos()
On Error Resume Next
Dim X As Long, Rs As Recordset
Set Rs = Banco.OpenRecordset("Lego", dbOpenDynaset)
For X = 1 To FrmCodigo.Eventos.Nodes.Count
    Rs.AddNew
    Rs!Key = FrmCodigo.Eventos.Nodes(X).Key
    Rs!Codigo = FrmCodigo.Eventos.Nodes(X).Tag
    Rs!Texto = FrmCodigo.Eventos.Nodes(X).Text
    Rs!Root = FrmCodigo.Eventos.Nodes(X).Root
    Rs.Update
Next X
Banco.Close
End Sub
Public Function BuscaImg(Img) As String
Dim Texto As String
Texto = ""
If Img <> 0 Then
    SavePicture Img, "C:\Cleber.INF.NEO.CAS"
    Texto = String(FileLen("C:\Cleber.INF.NEO.CAS"), " ")
    Open "C:\Cleber.INF.NEO.CAS" For Binary As #1
        Get #1, , Texto
    Close #1
    Kill "C:\Cleber.INF.NEO.CAS"
End If
BuscaImg = Texto
End Function
    
Public Function AbreFig(Texto As String, Obj As Object, Optional Tipo As Boolean)
If Tipo = True Then
    Obj.Picture = LoadPicture("")
Else
    Obj.Icon = LoadPicture("")
End If
If IsNull(Texto) = False Then
    If Trim(Texto) <> "" Then
        Open "C:\Cleber.inf.0001.aki.soares" For Binary As #1
            Put #1, , Texto
        Close #1
        If Tipo = True Then
            Obj.Picture = LoadPicture("C:\Cleber.inf.0001.aki.soares")
        Else
            Obj.Icon = LoadPicture("C:\Cleber.inf.0001.aki.soares")
        End If
        Kill "C:\Cleber.inf.0001.aki.soares"
    End If
End If
End Function

Public Sub GetEditStatus(Txt As RichTextBox, Label1 As Label)
   Dim lLine As Long, lCol As Long
   Dim cCol As Long, lChar As Long, i As Long

   lChar = Txt.SelStart + 1

   ' Get the line number
   lLine = 1 + SendMessageLong(Txt.HWnd, EM_LINEFROMCHAR, _
           Txt.SelStart, 0&)

   ' Get the Character Position
   cCol = SendMessageLong(Txt.HWnd, EM_LINELENGTH, lChar - 1, 0&)

   i = SendMessageLong(Txt.HWnd, EM_LINEINDEX, lLine - 1, 0&)
   lCol = lChar - i
   ' Caption of Label1 is set to Cursor Position.
   ' This could also be a panel in a StatusBar.
   Label1.Caption = lLine & " , " & lCol
   
End Sub

'Para usar:
Public Sub VerComandos1(Rtb As RichTextBox)
On Error Resume Next
Dim Comandos() As String
ReDim Comandos(99) As String
Dim X As Long, Y As Long

Comandos(0) = "SE"
Comandos(1) = "FIM SE"
Comandos(2) = "LOOP"
Comandos(3) = "VAPARA"
Comandos(4) = "FIM DE PROGRAMA"
Comandos(5) = "FECHA TELA"
Comandos(6) = "(TECLA)"
Comandos(7) = "CHAMA"
Y = 1
For X = 0 To 7
Inicio:
    Y = InStr(Y, UCase(Rtb.Text), Comandos(X))
    If Y <> 0 Then
        Rtb.SelStart = Y - 1
        Rtb.SelLength = Len(Comandos(X))
        Rtb.SelColor = &H800000
        Y = Y + 1
        GoTo Inicio
    End If
    Y = 1
Next X
End Sub
