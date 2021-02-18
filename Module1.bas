Attribute VB_Name = "ModRum"
Public FrmTela() As New Form2
Public ContTela As Long
Public TelaAtual As Long
Public IndexObj As Long
Public NomeBanco As String
Public Run() As New FrmR
Public Frm5 As Object
Public Modulo As String
Public LinhaErro As String
Public ContRun As Long
Public FrmTelaRun As Form
Public TextoAux As String
Public Bancos() As DBase
Public Rs() As DRs
Public TabRec() As TipoRecordSet

Public PosicaoDoFrmPrincipal As Byte
Type IdErrosNames
    ComandoInvalido As Long
    LoopInvalido As Long
    TelaInvalida As Long
    ObjInvalido As Long
    PropriedadeInvalida As Long
    BancoErro As Long
End Type
Type DBase
    Nome As String
    Banco As Database
    Local As String
End Type
Type TipoRecordSet
    Nome As String
    Tabela As String
    Ordem As String
    Codicao As String
    BancoDb As String
End Type
Type DRs
    Nome As String
    Rs As Recordset
End Type
Public NomeErros As IdErrosNames
Public NBanco As Byte
Public NRs As Long
Public Function WinCom(nTela As Long, TelaInicio As Form)
On Error Resume Next
Dim X As Long, index As Long
Dim Imenu As MenuItem, Ob As Object
Dim NOb As Object

Set Frm5 = TelaInicio
Set FrmTelaRun = Run(nTela)
Modulo = FrmCodigo.Eventos.Nodes(1).Tag

If Trim(Modulo) <> "" Then
    Rodar Modulo
End If
With TelaInicio
    Set FrmTelaRun = Run(nTela)
    Run(nTela).Tag = .Tag
    FrmPrincipal.ListaProg.AddItem .Tag
    Run(nTela).Height = .Height
    Run(nTela).Width = .Width
    Run(nTela).Caption = .Caption
    Run(nTela).BackColor = .BackColor
    Run(nTela).WindowState = .WindowState
    Run(nTela).Icon = .Icon
    
    'Desab Fecha(.Cont.Caption), Max(.Cont.Caption), Min(.Cont.Caption), Run(nTela)
    Run(nTela).Botao.Caption = IIf(Fecha(.Cont.Caption) = True, "1", "0") + IIf(Max(.Cont.Caption) = True, "1", "0") + IIf(Min(.Cont.Caption) = True, "1", "0")
    ', Run(nTela)
    
    Run(nTela).Xmenu.Refresh
'    Run(Ntela).ScaleHeight = .ScaleHeight
'    Run(Ntela).ScaleWidth = .ScaleWidth
    For Each Ob In TelaInicio
        If Ob.Name = "xMenu" Or Ob.Name = "Im" Or Ob.Name = "Focus" Or Ob.Name = "S" Or Ob.Name = "Nome" Or Ob.Name = "Cont" Then
            GoTo proximo
        ElseIf Ob.index = 0 Then
            GoTo proximo
        ElseIf Ob.Name = "Cmd" Then
            Load Run(nTela).Cmd(Run(nTela).Cmd.Count)
            Set NOb = Run(nTela).Cmd(Run(nTela).Cmd.Count - 1)
        ElseIf Ob.Name = "Fm" Then
            Load Run(nTela).Fm(Run(nTela).Fm.Count)
            Set NOb = Run(nTela).Fm(Run(nTela).Fm.Count - 1)
        ElseIf Ob.Name = "Img" Then
            Load Run(nTela).Img(Run(nTela).Img.Count)
            Set NOb = Run(nTela).Img(Run(nTela).Img.Count - 1)
        ElseIf Ob.Name = "Lbl" Then
            Load Run(nTela).Lbl(Run(nTela).Lbl.Count)
            Set NOb = Run(nTela).Lbl(Run(nTela).Lbl.Count - 1)
        ElseIf Ob.Name = "Chk" Then
            Load Run(nTela).Chk(Run(nTela).Chk.Count)
            Set NOb = Run(nTela).Chk(Run(nTela).Chk.Count - 1)
        ElseIf Ob.Name = "Cbo" Then
            Load Run(nTela).Cbo(Run(nTela).Cbo.Count)
            Set NOb = Run(nTela).Cbo(Run(nTela).Cbo.Count - 1)
        ElseIf Ob.Name = "Txt" Then
            Load Run(nTela).Txt(Run(nTela).Txt.Count)
            Set NOb = Run(nTela).Txt(Run(nTela).Txt.Count - 1)
        ElseIf Ob.Name = "Lst" Then
            Load Run(nTela).Lst(Run(nTela).Lst.Count)
            Set NOb = Run(nTela).Lst(Run(nTela).Lst.Count - 1)
        End If
        NOb.Appearance = Ob.Appearance
        NOb.Text = Ob.Text
        NOb.Stretch = Ob.Stretch
        NOb.BorderStyle = Ob.BorderStyle
        NOb.Caption = Ob.Caption
        NOb.Height = Ob.Height
        NOb.Width = Ob.Width
        NOb.BackColor = Ob.BackColor
        NOb.Tag = Ob.Tag
        NOb.Top = Ob.Top
        NOb.Left = Ob.Left
        NOb.Font = Ob.Font
        NOb.FontSize = Ob.FontSize
        NOb.FontName = Ob.FontName
        NOb.ForeColor = Ob.ForeColor
        NOb.ToolTipText = Ob.ToolTipText
        NOb.PasswordChar = Ob.PasswordChar
        NOb.Picture = Ob.Picture
        NOb.ZOrder = 1
        NOb.Visible = True
proximo:
        
    Next
    For X = 1 To .Cmd.Count - 1
        Run(nTela).Cmd(X).TabIndex = .Cmd(X).TabIndex
    Next X
    For X = 1 To .Fm.Count - 1
        Run(nTela).Fm(X).TabIndex = .Fm(X).TabIndex
    Next X
    For X = 1 To .Lbl.Count - 1
        Run(nTela).Lbl(X).TabIndex = .Lbl(X).TabIndex
    Next X
    For X = 1 To .Chk.Count - 1
        Run(nTela).Chk(X).TabIndex = .Chk(X).TabIndex
    Next X
    For X = 1 To .Txt.Count - 1
        Run(nTela).Txt(X).TabIndex = .Txt(X).TabIndex
    Next X
    For X = 1 To .Lst.Count - 1
        Run(nTela).Lst(X).TabIndex = .Lst(X).TabIndex
    Next X
    FrmTelaRun.Refresh
    Set FrmTelaRun = Run(nTela)
    Set Imenu = New MenuItem
    For X = 1 To Frm5.Xmenu.MenuTree.Count
        Imenu.Caption = Frm5.Xmenu.MenuTree(X).Caption
        Imenu.Name = Frm5.Xmenu.MenuTree(X).Name
        Imenu.Ident = Frm5.Xmenu.MenuTree(X).Ident

        FrmTelaRun.Xmenu.MenuTree.Add Imenu
        Set Imenu = New MenuItem
    Next X
    FrmTelaRun.Xmenu.Refresh
    FrmTelaRun.Xmenu.HighLightStyle = Frm5.Xmenu.HighLightStyle
    FrmTelaRun.Xmenu.ItemsFont = Frm5.Xmenu.ItemsFont
    FrmTelaRun.Xmenu.Style = Frm5.Xmenu.Style
    
    Rodar BuscaEventos(.Tag, "Carregar", "")
    
    Run(nTela).Visible = True
    Run(nTela).Xmenu.Refresh

End With

End Function

Public Function BuscaEventos(Tela As String, Obj As String, Optional Eventos As String)
Dim X As Long
For X = 1 To FrmCodigo.Eventos.Nodes.Count
    If Trim(Eventos) = "" Then
        If UCase(Tela) + "." + UCase(Obj) = UCase(FrmCodigo.Eventos.Nodes(X).Key) Then
            BuscaEventos = OrganizaSe(FrmCodigo.Eventos.Nodes(X).Tag)
            Exit Function
        End If
    Else
        If UCase(Tela) + "." + UCase(Obj) + "." + UCase(Eventos) = UCase(FrmCodigo.Eventos.Nodes(X).Key) Then
            BuscaEventos = OrganizaSe(FrmCodigo.Eventos.Nodes(X).Tag)
            Exit Function
        End If
    End If
Next X
End Function

Public Function Rodar(Texto As String)
On Error Resume Next
Dim Linha As String, Coluna As Long, Atual As Long
Dim SeDentro As Double
Dim ContSe As Long
Dim Entrei As String

Coluna = 1
Atual = 0
If Trim(Texto) = "" Then Exit Function

TextoAux = Texto + Chr(13) + Chr(10)
SeDentro = 0
Entrei = "-1"
Do While True
i:
    TextoAux = Trim(TextoAux)
    Coluna = InStr(1, TextoAux, Chr(10))
    If Coluna = 0 Or TextoAux = Chr(10) Or TextoAux = Chr(13) Then
        Exit Do
    End If
    Linha = Trim(Mid(TextoAux, Atual + 1, Coluna))
    LinhaErro = Trim(Linha)
    If Left(Linha, 1) = ";" Then
        buscaObj Linha
        TextoAux = Right(TextoAux, Len(TextoAux) - Coluna)
        GoTo i
    ElseIf UCase(Left(Trim(TextoAux), 8)) = "VA PARA " Then
        Linha = Trim(Right(Linha, Len(Linha) - 8))
        If Right(Linha, 2) = Chr(13) + Chr(10) Then
            Linha = Left(Linha, Len(Linha) - 2)
        End If
        Coluna = InStr(1, UCase(Texto), "LOOP " + UCase(Linha))
        If Coluna = 0 Then
            Erros NomeErros.LoopInvalido
            Exit Function
        End If
        TextoAux = Right(Texto, Len(Texto) - Coluna + 1) + Chr(13) + Chr(10)
        GoTo i
    ElseIf UCase(Left(TextoAux, 3)) = "SE:" Then
        Linha = Left(TextoAux, Coluna)
        Linha = Trim(Right(Linha, Len(Linha) - 2))
        Linha = Trim(Left(Linha, Len(Linha) - 2))
        SeDentro = SeDentro + 1
        Coluna1 = InStr(1, Linha, "=")
        p = Len(TextoAux)
        If Coluna1 <> 0 Then
            Dim Par1 As String, Par2 As String, Par3 As String
            Par1 = Left(Linha, Coluna1 - 1)
            p = InStr(1, Linha, " ")
            If p = 0 Then
                p = InStr(1, Linha, "=")
            End If
            If p = 0 Then
                'err
            End If
            Par3 = Trim(Mid(Linha, 1, p))
            Par2 = Trim(Right(Linha, Len(Linha) - Coluna1))
            Par1 = Trim(Right(Par1, Len(Par1) - p))
            
            Par1 = UCase(BuscaVar(Par1))
            Par2 = UCase(BuscaVar(Par2))
            Entrei = Par3
            If Par1 <> Par2 Then
                'Entrei = "-1"
                Par1 = Len("SeNao:" + Par3)
                Par2 = InStr(1, UCase(TextoAux), "SENAO" + Par3)
                If Par2 = 0 Then
                    Par1 = Len("FIMSE:" + Par3)
                    Par2 = InStr(1, UCase(TextoAux), "FIMSE" + Par3)
                End If
                If Par2 = 0 Then
                    'errr
                End If
                TextoAux = Right(TextoAux, Len(TextoAux) - Val(Par2) - Val(Par1))
            Else
                TextoAux = Trim(Right(TextoAux, Len(TextoAux) - Coluna))
            End If
            
        End If
'        Coluna1 = InStr(1, Linha, "<>")
'        If Coluna1 = 0 Then
'        Linha = Right(TextoAux, Len(TextoAux) - Coluna)
'        MsgBox Linha
        GoTo i
    ElseIf UCase(Left(TextoAux, 6)) = "SENAO:" Then
        If Entrei <> "-1" Then
            Par1 = Len("FIMSE" + Entrei)
            Par2 = InStr(1, UCase(TextoAux), "FIMSE" + Entrei)
            If Par2 = 0 Then
                Exit Function
                'erro
            End If
            TextoAux = Right(TextoAux, Len(TextoAux) - Val(Par2) + 1)
            GoTo i
        End If
    ElseIf UCase(Left(TextoAux, 14)) = "FAÇA ENQUANTO " Then
        Par1 = Right(TextoAux, Len(TextoAux) - 14)
        
    'elseif ucase(left(textoaux,10)) = "
    
    End If
    If Linha <> Chr(13) + Chr(10) And Asc(Linha) <> 10 Then
        If Trim(Linha) <> "" Then
            Linha = Trim(Linha)
            If Left(Linha, 1) <> "*" And Left(Linha, 1) <> "'" And Left(Linha, 2) <> "//" And Left(Linha, 2) <> "\\" Then
                LegoCom Linha
            End If
        End If
    End If
    If Trim(TextoAux) = "" Then Exit Do
    TextoAux = Trim(Right(TextoAux, Len(TextoAux) - Coluna))
Loop
End Function
Public Function LegoCom(Texto As String)
On Error Resume Next
Dim Part1 As String, Part2 As String, Coluna As String, Part3 As String
Dim X As Long
If UCase(Left(Texto, 4)) = "MSG " Then
    Part1 = Trim(Right(Texto, Len(Texto) - 4))
    Part1 = Left(Part1, Len(Part1) - 2)
    MsgBox VerVar(Part1), vbInformation, Titulo_Pjt
ElseIf UCase(Left(Texto, 5)) = "LOOP " Then

ElseIf UCase(Left(Texto, 13)) = "PROCEDIMENTO " Then
    If Right(Texto, 1) = Chr(10) Or Right(Texto, 1) = Chr(13) Then
        Texto = Left(Texto, Len(Texto) - 2)
    End If
    Rodar BuscaEventos("PROD", Trim(Right(Texto, Len(Texto) - 13)))

ElseIf UCase(Left(Texto, 15)) = "FIM DO PROGRAMA" Then
    For X = 0 To 100
        Unload Run(X)
    Next X
    TextoAux = ""
    If CRun.OpenExe = 0 Then
        End
    Else
        FrmPrincipal.WindowState = PosicaoDoFrmPrincipal
    End If
    
ElseIf UCase(Left(Texto, 13)) = "LIMPAARQUIVO(" Or UCase(Left(Texto, 14)) = "LIMPAARQUIVO (" Then
    Texto = VerrificaChr(Texto)
    If UCase(Left(Texto, 13)) = "LIMPAARQUIVO(" Then
        Texto = Right(Texto, Len(Texto) - 13)
    ElseIf UCase(Left(Texto, 14)) = "LIMPAARQUIVO (" Then
        Texto = Right(Texto, Len(Texto) - 13)
    End If
    Texto = Trim(Texto)
    Texto = Left(Texto, Len(Texto) - 1)
    Texto = VerVar(Texto)
    Open Texto For Output As #1
    Close #1

ElseIf UCase(Left(Texto, 8)) = "DELETAR(" Or UCase(Left(Texto, 9)) = "DELETAR (" Then
    Texto = VerrificaChr(Texto)
    If UCase(Left(Texto, 8)) = "DELETAR(" Then
        Texto = Right(Texto, Len(Texto) - 8)
    ElseIf UCase(Left(Texto, 9)) = "DELETAR (" Then
        Texto = Right(Texto, Len(Texto) - 9)
    End If
    Texto = Trim(Texto)
    Texto = Left(Texto, Len(Texto) - 1)
    Texto = VerVar(Texto)
    If Dir(Texto) = "" Or Texto = "" Then
        MsgBox "Impossivel apagar o arquivo , pois o caminho esta invalido", vbInformation, App.Title
    Else
        Kill Texto
    End If
ElseIf UCase(Left(Texto, 6)) = "GRAVA(" Or UCase(Left(Texto, 7)) = "GRAVA (" Then
    Texto = VerrificaChr(Texto)
    If UCase(Left(Texto, 6)) = "GRAVA(" Then
        Texto = Right(Texto, Len(Texto) - 6)
    ElseIf UCase(Left(Texto, 14)) = "GRAVA (" Then
        Texto = Right(Texto, Len(Texto) - 7)
    End If
    Texto = Trim(Texto)
    Texto = Left(Texto, Len(Texto) - 1)
    At = InStr(1, Texto, ",")
    If At = 0 Then
        MsgBox "Impossivel Grava no Arquivo !!!", vbInformation, App.Title
        Exit Function
    End If
    Part1 = Trim(Left(Texto, At - 1))
    Part2 = Trim(Right(Texto, Len(Texto) - At))
    Part1 = VerVar(Part1)
    Part2 = VerVar(Part2)
    Open Part1 For Binary As #1
        Put #1, , Part2
    Close #1


ElseIf UCase(Left(Texto, 3)) = "DB " Then
    AbreBanco Texto
ElseIf UCase(Left(Texto, 7)) = "TABELA " Then
    Tabelas Texto
ElseIf UCase(Left(Texto, 6)) = "FIMSE:" Then

ElseIf UCase(Left(Texto, 6)) = "SENAO:" Then
ElseIf UCase(Left(Texto, 6)) = "CHAMA " Then
    Part1 = Trim(Right(Texto, Len(Texto) - 6))
    Start Part1
ElseIf UCase(Left(Texto, 10)) = "FECHA TELA" Then
    Unload FrmTelaRun
ElseIf UCase(Left(Texto, 5)) = "TAB [" Then
    Tabelas Texto
ElseIf UCase(Left(Texto, 4)) = "EXE " Then
    Part1 = Trim(Right(Texto, Len(Texto) - 4))
    Shell BuscaVar(Part1), vbNormalFocus
ElseIf UCase(Left(Texto, 12)) = "TELA CHEIA()" Then
    FrmTelaRun.Top = 0
    FrmTelaRun.Left = 0
    'FrmTelaRun.Height = Screen.Height
    'FrmTelaRun.Width = Screen.Width
    FrmTelaRun.WindowState = 2
Else
    Coluna = InStr(1, Texto, "=")
    If Coluna = 0 Then
        If InStr(1, Texto, ".") <> 0 Then
        End If
        Erros NomeErros.ComandoInvalido
        Exit Function
    End If
    Part1 = Trim(Left(Texto, Coluna - 1))
    Part2 = Trim(Right(Texto, Len(Texto) - Coluna))
    Part2 = Left(Part2, Len(Part2) - 2)
    If Trim(Left(Part1, 1)) <> "*" Then
        Part2 = VerVar(Part2)
        EscrevaVar "Modulo", Part1, Part2
    End If
End If
End Function

Public Function BuscaVar(Part1 As String)

Dim Part2 As String
Dim Part3 As String
Dim Rest As Double

If UCase(Part1) = "(TECLA)" Then
    BuscaVar = Trim(Str(TeclaKey))
    Exit Function
End If
If Right(Part1, 2) = Chr(13) + Chr(10) Then
    Part1 = Trim(Left(Part1, Len(Part1) - 2))
End If
If Left(Part1, 1) = Chr(34) And Right(Part1, 1) = Chr(34) Then
    Part2 = Mid(Part1, 2, Len(Part1) - 2)

ElseIf IsNumeric(Part1) = True Then
    Part2 = Part1
Else
    If Left(Part1, 1) = ";" Then
        Part3 = VerVar(Part1)
        Part1 = Part3
        'Part3 = "=|="
    ElseIf Left(Part1, 1) = "@" Then
        Part3 = BuscaBanco(Part1)
    Else
        Part3 = LerVar("Modulo", Part1, "=|=")
    End If
    
    If Part3 = "=|=" Then
        If IsNumeric(Part1) = False Then
            If UCase(Left(Part1, 8)) = "SIMNAO (" Or UCase(Left(Part1, 7)) = "SIMNAO(" Then
                Part2 = Mid(Part1, 7, Len(Part1) - 1)
                If Left(Part2, 1) = "(" Then
                    Part2 = Right(Part2, Len(Part2) - 1)
                End If
                If Right(Part2, 1) = ")" Then
                    Part2 = Left(Part2, Len(Part2) - 1)
                End If
                Part2 = BuscaVar(Part2)
                Rest = MsgBox(Part2, vbInformation + vbYesNo + vbSystemModal, Titulo_Pjt)
                If Rest = vbYes Then
                    Part3 = "Sim"
                Else
                    Part3 = "Não"
                End If
                BuscaVar = Part3
                Exit Function
            ElseIf UCase(Left(Part1, 8)) = "DIREITA(" Or UCase(Left(Part1, 9)) = "DIREITA (" Then
                If UCase(Left(Part1, 8)) = "DIREITA(" Then
                    Part3 = Right(Part1, Len(Part1) - 8)
                ElseIf UCase(Left(Part1, 9)) = "DIREITA (" Then
                    Part3 = Right(Part1, Len(Part1) - 9)
                End If
                Part3 = Left(Part3, Len(Part3) - 1)
                At = InStr(1, Part3, ",")
                If At = 0 Then
                    Exit Function
                End If
                Part2 = BuscaVar(Left(Part3, At - 1))
                Part3 = BuscaVar(Right(Part3, Len(Part3) - At))
                Part2 = Right(Part2, Part3)
                BuscaVar = Part2
                Exit Function
            ElseIf UCase(Left(Part1, 9)) = "ESQUERDA(" Or UCase(Left(Part1, 10)) = "ESQUERDA (" Then
                Part1 = Trim(Part1)
                If UCase(Left(Part1, 9)) = "ESQUERDA(" Then
                    Part3 = Right(Part1, Len(Part1) - 9)
                ElseIf UCase(Left(Part1, 10)) = "ESQUERDA (" Then
                    Part3 = Right(Part1, Len(Part1) - 10)
                End If
                Part3 = Left(Part3, Len(Part3) - 1)
                At = InStr(1, Part3, ",")
                If At = 0 Then
                    Exit Function
                End If
                Part2 = BuscaVar(Left(Part3, At - 1))
                Part3 = BuscaVar(Right(Part3, Len(Part3) - At))
                Part2 = Left(Part2, Part3)
                BuscaVar = Part2
                Exit Function
            ElseIf UCase(Left(Part1, 6)) = "CAIXA(" Or UCase(Left(Part1, 7)) = "CAIXA (" Then
                Part1 = Trim(Part1)
                If UCase(Left(Part1, 6)) = "CAIXA(" Then
                    Part3 = Right(Part1, Len(Part1) - 6)
                ElseIf UCase(Left(Part1, 10)) = "CAIXA (" Then
                    Part3 = Right(Part1, Len(Part1) - 7)
                End If
                If Right(Part3, 1) <> ")" Then
                    'Colocar o erro aqui....
                End If
                Part3 = Trim(Left(Part3, Len(Part3) - 1))
                Part2 = Left(Part3, Len(Part3) - 1)
                Part2 = Trim(Part2)
                Part2 = Trim(Left(Part2, Len(Part2) - 1))
                Part3 = Trim(Part3)
                Part1 = Right(Part3, 1)
                Part2 = BuscaVar(Trim(Part2))
                FrmTelaRun.ComRun.Filter = Part2
                If Part1 = "0" Then
                    FrmTelaRun.ComRun.DialogTitle = Titulo_Pjt + " - Abrir"
                    FrmTelaRun.ComRun.ShowOpen
                Else
                    FrmTelaRun.ComRun.DialogTitle = Titulo_Pjt + " - Salvar"
                    FrmTelaRun.ComRun.ShowSave
                End If
                BuscaVar = FrmTelaRun.ComRun.FileName
                Exit Function
            ElseIf UCase(Left(Part1, 6)) = "ABRIR(" Or UCase(Left(Part1, 7)) = "ABRIR (" Then
                Part1 = Trim(Part1)
                If UCase(Left(Part1, 6)) = "ABRIR(" Then
                    Part3 = Right(Part1, Len(Part1) - 6)
                ElseIf UCase(Left(Part1, 10)) = "ABRIR (" Then
                    Part3 = Right(Part1, Len(Part1) - 7)
                End If
                If Right(Part3, 1) <> ")" Then
                    'Colocar o erro aqui....
                End If
                Part3 = Trim(Left(Part3, Len(Part3) - 1))
                Part3 = VerVar(Part3)
                If Dir(Part3) = "" Or Part3 = "" Then
                    MsgBox "Impossivel Abrir o arquivo" + Chr(13) + "Caminho Invalido ! ! !", vbInformation, App.Title
                    BuscaVar = ""
                    Exit Function
                End If
                Part2 = String(FileLen(Part3), " ")
                Open Part3 For Binary As #1
                    Get #1, , Part2
                Close #1
                BuscaVar = Part2
                Exit Function
            ElseIf Part1 <> "" Then
                If TestaVarNun(Part1) = True Then
                    Part1 = VerVar(Part1)
                End If
            End If
        End If
        If GetValue(ParseInit, Part1, Rest) Then
            Part3 = Trim(Str(Rest))
            Part2 = Trim(Str(Rest))
        Else
            If Len(Part1) > 1 Then
                If InStr(1, Part1, " + ") <> 0 Then
                    Part2 = Replace(Part1, " + ", "")
                ElseIf InStr(1, Part1, " +") <> 0 Then
                    Part2 = Replace(Part1, " +", "")
                ElseIf InStr(1, Part1, "+ ") <> 0 Then
                    Part2 = Replace(Part1, "+ ", "")
                ElseIf InStr(1, Part2, "+") <> 0 Then
                    Part2 = Replace(Part1, "+", "")
                End If
            Else
                Part2 = Part1
            End If
        End If
        
        
    Else
        Part2 = Part3
    End If
End If
BuscaVar = Part2
End Function

Public Function VerVar(Var As String)
On Error Resume Next
Dim X As Long, Inicio As Long
Dim NovaVar As String, Posicao() As String
Dim Xy As Long, Part As String
ReDim Posicao(11) As String
Dim Antiga As String
Dim AntigaVar As String
Dim Var1 As String
Dim Po1 As Long
Dim TipoString As Long
Dim DentroParente As Boolean

Xy = 0
AntigaVar = Var
Var1 = Var
NovaVar = Var
Inicio1 = 1
Antiga = ""
Var = VerrificaChr(Var)
Var = Var + " "
Dentro = True
Po1 = 0
TipoString = 0
DentroParente = True


For Xy = 1 To Len(Var)
    Var1 = Asc(Mid(Var, Xy, 1))
    X = Asc(UCase(Mid(Var, Xy, 1)))
    If X = Asc(";") Then
        Po1 = Xy
        TipoString = 1
        Antiga = Antiga + Chr(Var1)
    ElseIf X = Asc("(") Then
        Po1 = Xy - Len(Antiga)
        Dentro = False
        TipoString = 2
        Antiga = Antiga + Chr(X)
        DentroParente = False
    ElseIf X = Asc(")") Then
        Antiga = Antiga + Chr(X)
        Dentro = True
        DentroParente = True
    ElseIf X >= vbKeyA - 1 And X <= vbKeyZ Or X = Asc(";") Or X = Asc(".") Then
        Antiga = Antiga + Chr(Var1)
    ElseIf X >= vbKey0 And X <= vbKey9 Then
        Antiga = Antiga + Chr(Var1)
    ElseIf X = 34 Then
        Antiga = Antiga + Chr(Var1)
        Dentro = Not Dentro
    Else
        If Antiga <> "" Then
            If Dentro = True And DentroParente = True Then
                At = Antiga
                If TipoString = 1 Then
                    AntigaVar = BuscaObjto(Antiga)
                Else
                    AntigaVar = BuscaVar(Antiga)
                End If
              '  If Left(Antiga, 1) = Chr(34) And Right(Antiga, 1) = Chr(34) Then
              
                NovaVar = Replace(NovaVar, At, AntigaVar)
               ' End If
               TipoString = 0
                Antiga = ""
            Else
                Antiga = Antiga + Chr(Var1)
            End If
        End If
    End If
    DoEvents
Next

NovaVar = Trim(NovaVar)
If TestaVarNun(NovaVar) = True Then
    If GetValue(0, NovaVar, Rest) Then
        Part3 = Trim(Str(Rest))
        Part2 = Trim(Str(Rest))
        VerVar = Part2
    End If
Else
    VerVar = NovaVar
End If
End Function

Public Sub Start(Nometela As String)
Dim X As Long, Passa As Boolean

Passa = True
NomeErros.ComandoInvalido = 0
NomeErros.LoopInvalido = 1
NomeErros.TelaInvalida = 2
NomeErros.ObjInvalido = 3
NomeErros.PropriedadeInvalida = 4
NomeErros.BancoErro = 5
NBanco = 0
NRs = 0
If Right(Nometela, 2) = Chr(13) + Chr(10) Then
    Nometela = Left(Nometela, Len(Nometela) - 2)
End If
For X = 0 To ContTela
    If UCase(FrmTela(X).Tag) = UCase(Nometela) Then
        Passa = False
        Exit For
    End If
Next X
If Passa = True Then
    Erros NomeErros.TelaInvalida
    Exit Sub
End If
Abilita False, True
WinCom ContRun, FrmTela(X)
ContRun = ContRun + 1
End Sub


Private Sub buscaObj(Par As String)
On Error Resume Next
Dim Linha As String, L As Long, Par2 As String
Dim X As Long, TelaPrincipal As Long
Dim Obj As Object, Nometela As String, NomeObj As String
Dim Ob As Object, Propria As String, L1 As Long

Linha = Trim(Par)
Linha = Right(Linha, Len(Linha) - 1)
L = InStr(1, Linha, "=")
L1 = L

If L = 0 Then
    GoTo Passa1
    Exit Sub
End If

Par2 = Right(Linha, Len(Linha) - L)
Linha = Left(Linha, L - 1)

Passa1:

L = 1
L = InStr(L, Linha, ".")
If L = 0 Then
    Nometela = UCase(FrmTelaRun.Tag)
    Propria = Trim(UCase(Linha))
    Set Obj = FrmTelaRun
    GoTo Proximo1
Else
    Nometela = UCase(Left(Linha, L - 1))
End If
For X = 0 To ContTela
    If Nometela = UCase(Run(X).Tag) Then
        TelaPrincipal = X
        Set Obj = Run(X)
        GoTo Passa
    End If
Next

Erros:
L = InStr(L, Linha, ".")
If L = 0 Then GoTo Erros
NomeObj = UCase(Left(Linha, L - 1))
a = InStr(NomeObj, ".")
If a <> 0 Then
    NomeObj = Left(NomeObj, a - 1)
    L = L + a
End If
For Each Ob In FrmTelaRun
    If UCase(Ob.Tag) = NomeObj Then
        Set Obj = Ob
        GoTo proximo
    End If
Next
'errr
GoTo proximo

Exit Sub

Passa:
    L = InStr(L, Linha, ".")
    If L = 0 Then GoTo Erros
    NomeObj = UCase(Mid(Linha, L + 1, Len(Linha) - L))
    a = InStr(NomeObj, ".")
    If a <> 0 Then
        NomeObj = Left(NomeObj, a - 1)
        L = L + a
    End If
    For Each Ob In Run(TelaPrincipal)
        If UCase(Ob.Tag) = NomeObj Then
            Set Obj = Ob
            GoTo proximo
        End If
    Next
'erro

proximo:
Propria = Trim(UCase(Right(Linha, Len(Linha) - L)))

Proximo1:
If L1 <> 0 Then
    Par2 = Trim(Par2)
    Par2 = VerVar(Par2)
    If Right(Par2, 2) = Chr(13) + Chr(10) Then
        Par2 = Left(Par2, Len(Par2) - 2)
    End If
Else
    If Right(Propria, 2) = Chr(13) + Chr(10) Then
        Propria = Left(Propria, Len(Propria) - 2)
    End If
End If
On Error GoTo Trata_Erro
Select Case Propria
    
    Case "LEGENDA"
        Obj.Caption = Par2
    Case "COR DE FUNDO", "COR FUNDO"
        Obj.BackColor = Par2
    Case "COR DA LETRA", "COR LETRA"
        Obj.ForeColor = Par2
    Case "TAMX"
        Obj.Height = Par2
    Case "TAMY"
        Obj.Width = Par2
    Case "POXX"
        Obj.Top = Par2
    Case "POXY"
        Obj.Left = Par2
    Case "FONTE"
        Obj.FontName = Par2
    Case "TAMANHO"
        Obj.FontSize = Par2
    Case "TEXTO"
        Obj.Text = Par2
    Case "IMAGEM"
        If Dir(Par2) = "" Then
            MsgBox "Imagem Invalida ! ! !", vbCritical, App.Title
            Exit Sub
        End If
        Obj.ToolTipText = Par2
        Obj.Picture = LoadPicture(Par2)
    Case "BORDA"
        If Left(Par2, 1) = "S" Then
            Par2 = "Sim"
            Obj.BorderStyle = 1
        Else
            Obj.BorderStyle = 0
            Par2 = "Não"
        End If
    Case "COMPRIMIR"
        If Left(Par2, 1) = "S" Then
            Par2 = "Sim"
            Obj.Stretch = 1
        Else
            Obj.Stretch = 0
            Par2 = "Não"
        End If
     
    Case "ORDEM"
        If IsNumeric(Par2) = True Then
            Obj.TabIndex = Par2
        Else
            Par2 = Obj.TabIndex
        End If
    Case "MARCA"
        If Left(Par2, 1) = "S" Then
            Par2 = "Sim"
            Obj.Value = 1
        Else
            Obj.Value = 0
            Par2 = "Não"
        End If
    Case "FOCUS"
        Obj.SetFocus
    Case "ADICIONE"
        Obj.AddItem Par2
    Case "SELECIONE"
        Obj.Selected(Par2) = True
    Case "LIMPA"
        Obj.Clear
    Case "VISIVEL"
        If Left(Par2, 1) = "S" Then
            Par2 = "Sim"
            Obj.Visible = True
        Else
            Obj.Visible = False
            Par2 = "Não"
        End If
    Case Else
        Erros NomeErros.PropriedadeInvalida
        Exit Sub
End Select
Obj.Refresh
Exit Sub

Trata_Erro:
Erros NomeErros.PropriedadeInvalida
End Sub


Public Function BuscaObjto(Nome As String)
On Error GoTo Trata_Erro
Dim Linha As String, L As Long, Par2 As String
Dim X As Long, TelaPrincipal As Long
Dim Obj As Object, Nometela As String, NomeObj As String
Dim Ob As Object, Propria As String
Dim Retorno As String

Retorno = "Erro , Objeto Não Indetificado"
Linha = Trim(Replace(Nome, ";", " "))

'Par2 = Right(Linha, Len(Linha) - L)
'Linha = Left(Linha, L - 1)

L = 1
L = InStr(L, Linha, ".")
If L = 0 Then
    Nometela = UCase(FrmTelaRun.Tag)
    Propria = Trim(UCase(Linha))
    Set Obj = FrmTelaRun
    GoTo Proximo1
Else
    Nometela = UCase(Left(Linha, L - 1))
End If
For X = 0 To ContTela - 1
    If Nometela = UCase(FrmTela(X).Tag) Then
        TelaPrincipal = X
        Set Obj = Run(X)
        GoTo Passa
    End If
Next

Erros:
'erro
NomeObj = Nometela
Set Obj = FrmTelaRun

GoTo Pri
Exit Function

Passa:
    L = InStr(L, Linha, ".")
    If L = 0 Then GoTo Erros
    NomeObj = UCase(Mid(Linha, L + 1, Len(Linha) - L))
    a = InStr(NomeObj, ".")
    If a <> 0 Then
        NomeObj = Left(NomeObj, a - 1)
        L = L + a
    End If
Pri:
    For Each Ob In Obj
        If UCase(Ob.Tag) = NomeObj Then
            Set Obj = Ob
            GoTo proximo
        End If
    Next
'erro

proximo:
Proximo1:
Propria = Trim(UCase(Right(Linha, Len(Linha) - L)))

Select Case Propria
    
    Case "LEGENDA"
        Retorno = Obj.Caption
    Case "COR DE FUNDO", "COR FUNDO"
       Retorno = Obj.BackColor
    Case "COR DA LETRA", "COR LETRA"
        Retorno = Obj.ForeColor
    Case "TAMX"
        Retorno = Obj.Height
    Case "TAMY"
        Retorno = Obj.Width
    Case "POXX"
        Retorno = Obj.Top
    Case "POXY"
        Retorno = Obj.Left
    Case "FONTE"
        Retorno = Obj.FontName
    Case "TAMANHO"
        Retorno = Obj.FontSize
    Case "TEXTO"
        Retorno = Obj.Text
    Case "IMAGEM"
        Retorno = Obj.ToolTipText
    Case "BORDA"
        If Obj.BorderStyle = 1 Then
            Par2 = "Sim"
        Else
            Par2 = "Não"
        End If
        Retorno = Par2
    Case "COMPRIMIR"
        If Obj.Stretch = 1 Then
            Par2 = "Sim"
        Else
            Par2 = "Não"
        End If
        Retorno = Par2
    Case "ORDEM"
        Retorno = Obj.TabIndex
    Case "3D"
        If Obj.Appearance = 1 Then
            Par2 = "Sim"
        Else
            Par2 = "Não"
        End If
        Retorno = Par2
    Case "MARCA"
        If Obj.Value = 1 Then
            Par2 = "Sim"
        Else
            Par2 = "Não"
        End If
        Retorno = Par2
    Case "MASCARA"
        Retorno = Obj.PasswordChar
    Case "TOTAL"
        Retorno = Obj.ListCount
    Case "ATUAL"
        Retorno = Obj.ListIndex
    Case "BUSCA"
        Retorno = Obj.List(Par2)
    Case "VISIVEL"
        If Obj.Visible = True Then
            Retorno = "Sim"
        Else
            Retorno = "Não"
        End If
    Case Else
        Erros NomeErros.PropriedadeInvalida
        
End Select
BuscaObjto = Retorno
Exit Function

Trata_Erro:
BuscaObjto = ""
End Function


Private Function Erros(Id As Long)  ', Optional Comentarios As String)
On Error Resume Next
Dim Nome As String

If Err.Number = 13 Then Exit Function

If Id = NomeErros.ComandoInvalido Then
    Nome = LinhaErro + Chr(13) + Chr(13) + "Comando Desconhecido" + Chr(13) + "N.º 1285"
ElseIf Id = NomeErros.LoopInvalido Then
    Nome = LinhaErro + Chr(13) + Chr(13) + "Comando Desconhecido" + Chr(13) + "N.º 125" + Chr(13) + "Incio do loop não encontrado"
ElseIf Id = NomeErros.TelaInvalida Then
    Nome = "Não foi Possivel Encontrar a Tela " + Nometela + Chr(13)
    Nome = Nome + "N.º 0101 - Descricao Invalida"
ElseIf Id = NomeErros.PropriedadeInvalida Then
    Nome = "Propriedade do Objeto não Encoda " + Chr(13) + LinhaErro + Chr(13) + "N.º 1001"
ElseIf Id = NomeErros.BancoErro Then
    Nome = "Banco de dados invalido , ou inexistente" + Chr(13) + Chr(13) + "N.º 0023"
ElseIf Id = 6 Then
    Nome = "Impossivel Abrir a Tabela "
End If
Nome = Nome + Chr(13) + Chr(13) + LinhaErro
FrmErro.Erro.Caption = Nome
FrmErro.Show 1
TextoAux = ""
FrmPrincipal.WindowState = 2
End Function

Public Function OrganizaSe(Comandos As String)
Attribute OrganizaSe.VB_UserMemId = 0
On Error Resume Next
Dim NovoComando As String
Dim X As Long, ContSe As Long
Dim Nov As String
Dim TextoA As String
Dim NovoTexto As String
Dim TextoN As String
Dim Guarda As Long
Dim LenLinha As Long
X = 0
If Trim(Comandos) = "" Then Exit Function
ContSe = 1
NovoComando = Comandos
Nov = Comandos
NovoTexto = ""
Do While True
Inicio:
    X = InStr(1, NovoComando, Chr(10))
    If X = 0 Then
        If NovoComando = "" Then
            Exit Do
        Else
            X = Len(NovoComando)
        End If
    End If
    Guarda = X
    
    TextoA = Left(NovoComando, X)
    LenLinha = Len(TextoA)
    X = InStr(1, UCase(TextoA), "SE ")
    If X <> 0 Then
        TextoA = Left(TextoA, X + 1) + ":" + Trim(Str(ContSe)) + Right(TextoA, Len(TextoA) - X - 1)
        ContSe = ContSe + 1
        NovoTexto = NovoTexto + TextoA
        NovoComando = Right(NovoComando, Len(NovoComando) - LenLinha)
        GoTo Inicio
    End If
    
    'TextoA = Left(NovoComando, X + 1)
    X = InStr(1, UCase(TextoA), "SENAO")
    If X <> 0 Then
        TextoA = Left(TextoA, X + 4) + ":" + Trim(Str(ContSe - 1)) + Right(TextoA, Len(TextoA) - X - 4)
        NovoTexto = NovoTexto + TextoA
        NovoComando = Right(NovoComando, Len(NovoComando) - LenLinha)
        GoTo Inicio
    End If
    
    'TextoA = Left(NovoComando, X + 1)
    X = InStr(1, UCase(TextoA), "FIMSE")
    If X <> 0 Then
        TextoA = Left(TextoA, X + 4) + ":" + Trim(Str(ContSe - 1)) + Right(TextoA, Len(TextoA) - X - 4)
        NovoTexto = NovoTexto + TextoA
        NovoComando = Right(NovoComando, Len(NovoComando) - LenLinha)
        ContSe = ContSe - 1
        GoTo Inicio
    End If
    NovoTexto = NovoTexto + TextoA
    NovoComando = Right(NovoComando, Len(NovoComando) - LenLinha)
    ''NovoComando = Left(NovoComando, X + 1) + ":" + Trim(Str(ContSe)) + Right(NovoComando, Len(NovoComando) - X - 1)
Loop
OrganizaSe = NovoTexto
Open "C:\TESTE" For Output As #1
    Print #1, NovoTexto
Close #1
End Function

Public Sub AbreBanco(Texto As String)
Dim Par1 As String, Part2 As String
Dim Cont As Long
Texto = Trim(Right(Texto, Len(Texto) - 3))
Cont = InStr(1, Texto, ";")
If Cont = 0 Then
    Erros 5
    Exit Sub
End If
Part1 = BuscaVar(Trim(Left(Texto, Cont - 1)))
Part2 = Trim(Right(Texto, Len(Texto) - Cont))
If Right(Part2, 1) = Chr(10) Or Left(Part2, 1) = Chr(13) Then
    Part2 = Left(Part2, Len(Part2) - 2)
End If
If Dir(Part1) = "" Then
    Erros 5
    Exit Sub
End If
Set Bancos(NBanco).Banco = OpenDatabase(Part1)
Bancos(NBanco).Local = Part1
Bancos(NBanco).Nome = Part2
NBanco = NBanco + 1
End Sub

Private Function Tabelas(Texto As String)
On Error Resume Next
Dim Pesq As Long, Pesq1 As Long
Dim X As Long, Nome As String
Dim Banco1 As String
Dim Sql As String
Dim B As Database
Dim At As Long
Dim R As Recordset

If Right(Texto, 1) = Chr(10) Or Right(Texto, 1) = Chr(13) Then
    Texto = Left(Texto, Len(Texto) - 2)
End If
Texto = Trim(Right(Texto, Len(Texto) - 7))
Pesq = InStr(1, Texto, "=")
If Pesq = 0 Then
    Erros 6
    Exit Function
End If
Nome = Trim(Left(Texto, Pesq - 1))
Pesq1 = InStr(1, Texto, "[")

If Pesq1 = 0 Then
    Erros 6
    Exit Function
End If
Banco1 = Trim(Mid(Texto, Pesq + 1, Pesq1 - Pesq - 1))

Sql = Right(Texto, Len(Texto) - Pesq1)

Sql = Left(Sql, Len(Sql) - 1)
Sql = BuscaVar(Sql)
At = NRs
For X = 0 To NRs
    If UCase(Nome) = UCase(Rs(X).Nome) Then
        At = X
        GoTo Nao
    End If
Next
Cont:
NRs = NRs + 1

Nao:
For X = 0 To NBanco
    If UCase(Bancos(X).Nome) = UCase(Banco1) Then
        Set B = Bancos(X).Banco
        Exit For
    End If
Next X

Rs(At).Nome = Nome
'MsgBox Bancos(0).Banco(Sql).Fields(0).Value
Set R = Bancos(X).Banco.OpenRecordset(Sql)
Set Rs(At).Rs = R


If Err Then
    MsgBox Err.Description & "   " & Err.Number
End If
End Function


Private Function BuscaBanco(Texto)
On Error Resume Next
On Error Resume Next
Dim Pesq As Long, Pesq1 As Long
Dim X As Long, Nome As String
Dim Banco1 As String
Dim Sql As String
Dim B As Database
Dim At As Long


If Right(Texto, 1) = Chr(10) Or Right(Texto, 1) = Chr(13) Then
    Texto = Left(Texto, Len(Texto) - 2)
End If

Pesq = InStr(1, Texto, "{")

If Pesq = 0 Then
    Erros 6
    Exit Function
End If
Nome = Trim(Left(Texto, Pesq - 1))
Nome = Right(Nome, Len(Nome) - 1)

Sql = Right(Texto, Len(Texto) - Pesq)

Sql = Left(Sql, Len(Sql) - 1)
Sql = BuscaVar(Sql)
At = NRs
For X = 0 To NRs
    If UCase(Nome) = UCase(Rs(X).Nome) Then
        At = X
        Exit For
    End If
Next
BuscaBanco = Rs(At).Rs(Sql).Value

If Err Then
    MsgBox Err.Description & "   " & Err.Number
End If

End Function

Private Function TestaVarNun(Texto As String) As Boolean
On Error Resume Next
Dim X As Long

TestaVarNun = True
For X = 1 To Len(Texto)
    If IsNumeric(Mid(Texto, X, 1)) = False Then
        Select Case Mid(Texto, X, 1)
            Case "+", "-", "/", "*", " "
                
            Case Else
             TestaVarNun = False
             Exit For
        End Select
    End If
Next X
        

End Function

Public Function VerrificaChr(Texto As String)
On Error Resume Next
Dim X As Long
For X = 0 To 1000
    If Right(Texto, 1) = Chr(10) Or Right(Texto, 1) = Chr(13) Then
        Texto = Left(Texto, Len(Texto) - 1)
    Else
        Exit For
    End If
Next
VerrificaChr = Texto
End Function
Public Function Abilita(Tipo As Boolean, Optional Tipo1 As Boolean)
On Error Resume Next
Dim X As Long
With FrmPrincipal
    .MenuSalvar.Enabled = Tipo
    .MenuSalComo.Enabled = Tipo
    .MenuCompli.Enabled = Tipo
    .MenuProjeto.Enabled = Tipo
    .MenuPropriedade.Enabled = Tipo
    .MenuCod.Enabled = Tipo
    .MenuFerra.Enabled = Tipo
    .MenuEdito.Enabled = Tipo
    .MenuEnviarTraz.Enabled = Tipo
    .MenuEnviar.Enabled = Tipo
    .menuComplile.Enabled = Tipo
    .menuopcoes.Enabled = Tipo
    'MenuBancoDados.Enabled = Tipo
    .MenuTela.Enabled = Tipo
    .T.Enabled = Tipo
    .Toolbar1.Buttons(3).Enabled = Tipo
    .Toolbar1.Buttons(7).Enabled = Tipo
    .Toolbar1.Buttons(8).Enabled = Not Tipo
    .Toolbar1.Buttons(1).ButtonMenus(2).Enabled = Tipo
    'Toolbar1.Buttons(1).ButtonMenus(3).Enabled = Tipo
    'Toolbar1.Buttons(1).ButtonMenus(4).Enabled = Tipo
    .MenuLimpa.Enabled = Tipo
    .menuParaExecutar.Enabled = Not Tipo
    If Tipo1 = True Then
        .MenuAbrir.Enabled = Tipo
        .menuNovo.Enabled = Tipo
        .menuopcoes.Enabled = Tipo
        .MenuProcedimentoPublic.Enabled = Tipo
        .MenuLimpa.Enabled = Tipo
        .Toolbar1.Buttons(1).Enabled = Tipo
        .Toolbar1.Buttons(2).Enabled = Tipo
    End If
End With
'FrmCodigo.TxtCod.TxtCod.Locked = Not Tipo
'FrmPrincipal.. Enabled = Tipo
'FrmPrincipal.Enabled = Tipo
End Function
