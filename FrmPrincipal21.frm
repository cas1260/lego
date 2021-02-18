VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{A6BDE5D5-8F7A-11D1-9C65-4CA605C10E27}#5.0#0"; "ACTIVEGUI.OCX"
Begin VB.MDIForm FrmPrincipal 
   BackColor       =   &H00808000&
   Caption         =   "Lego 1.0"
   ClientHeight    =   5955
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6465
   Icon            =   "FrmPrincipal21.frx":0000
   LinkTopic       =   "MDIForm1"
   MousePointer    =   11  'Hourglass
   ScrollBars      =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin ActiveGUICtl.ActiveDock Ferramentas 
      Align           =   3  'Align Left
      Height          =   5595
      Left            =   0
      Negotiate       =   -1  'True
      TabIndex        =   1
      Top             =   360
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   9869
      Caption         =   "Ferramentas"
      Begin MSComctlLib.Toolbar T 
         Height          =   570
         Left            =   150
         TabIndex        =   2
         Top             =   480
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   1005
         ButtonWidth     =   926
         ButtonHeight    =   900
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   10
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "A"
               Object.ToolTipText     =   "Seleção"
               ImageIndex      =   1
               Style           =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "B"
               Object.ToolTipText     =   "Botão"
               ImageIndex      =   2
               Style           =   1
               Object.Width           =   1e-4
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "C"
               Object.ToolTipText     =   "Frame"
               ImageIndex      =   3
               Style           =   1
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "D"
               Object.ToolTipText     =   "Imagem"
               ImageIndex      =   4
               Style           =   1
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "E"
               Object.ToolTipText     =   "Palavras"
               ImageIndex      =   5
               Style           =   1
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "F"
               Object.ToolTipText     =   "Check"
               ImageIndex      =   6
               Style           =   1
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "G"
               Object.ToolTipText     =   "Texto"
               ImageIndex      =   7
               Style           =   1
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "H"
               Object.ToolTipText     =   "ComboBox"
               ImageIndex      =   8
               Style           =   1
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "J"
               Object.ToolTipText     =   "Lista Box"
               ImageIndex      =   9
               Style           =   1
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Key             =   "I"
               Object.ToolTipText     =   "Banco de Dados"
               ImageIndex      =   10
            EndProperty
         EndProperty
         Enabled         =   0   'False
      End
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   1530
      Top             =   3330
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrincipal21.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrincipal21.frx":0556
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrincipal21.frx":066A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrincipal21.frx":2CB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrincipal21.frx":5302
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrincipal21.frx":59D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrincipal21.frx":5F72
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrincipal21.frx":684E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog Com 
      Left            =   1770
      Top             =   2490
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3480
      Top             =   2790
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483624
      ImageWidth      =   28
      ImageHeight     =   28
      MaskColor       =   16776960
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrincipal21.frx":712A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrincipal21.frx":7AAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrincipal21.frx":8432
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrincipal21.frx":8DB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrincipal21.frx":973A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrincipal21.frx":A0BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrincipal21.frx":AA42
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrincipal21.frx":B3C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrincipal21.frx":BD4A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrincipal21.frx":C6CE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6465
      _ExtentX        =   11404
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "A"
            Object.ToolTipText     =   "Nova Tela"
            ImageIndex      =   1
            Style           =   5
            Object.Width           =   1e-4
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "NP"
                  Text            =   "Novo Projeto       "
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Key             =   "NT"
                  Text            =   "Nova Tela"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "B"
            Object.ToolTipText     =   "Abrir"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "C"
            Object.ToolTipText     =   "Salvar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "D"
            ImageIndex      =   4
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "E"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "F"
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "G"
            Object.ToolTipText     =   "Executar"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "H"
            Object.ToolTipText     =   "Para a Execução"
            ImageIndex      =   6
         EndProperty
      EndProperty
   End
   Begin VB.Menu MenuArquivo 
      Caption         =   "&Arquivo"
      Begin VB.Menu menuNovo 
         Caption         =   "&Novo"
         Begin VB.Menu MenuProjetog 
            Caption         =   "Projeto"
         End
         Begin VB.Menu MenuBranco 
            Caption         =   "-"
            Index           =   0
         End
         Begin VB.Menu MenuTela 
            Caption         =   "Tela"
            Enabled         =   0   'False
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
         Enabled         =   0   'False
      End
      Begin VB.Menu MenuSalComo 
         Caption         =   "Salvar &Como ..."
         Enabled         =   0   'False
      End
      Begin VB.Menu meudddd 
         Caption         =   "-"
      End
      Begin VB.Menu MenuCompli 
         Caption         =   "&Compile                                      "
         Enabled         =   0   'False
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
      Begin VB.Menu MenuProjeto 
         Caption         =   "&Projetos"
         Enabled         =   0   'False
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnusss 
         Caption         =   "-"
      End
      Begin VB.Menu MenuPropriedade 
         Caption         =   "Propriedade"
         Enabled         =   0   'False
         Shortcut        =   {F4}
      End
      Begin VB.Menu MenuBr 
         Caption         =   "-"
      End
      Begin VB.Menu MenuCod 
         Caption         =   "Codigos"
         Shortcut        =   {F7}
      End
      Begin VB.Menu menuSep 
         Caption         =   "-"
      End
      Begin VB.Menu MenuFerra 
         Caption         =   "Ferramentas"
         Checked         =   -1  'True
         Enabled         =   0   'False
         Shortcut        =   ^{F7}
      End
      Begin VB.Menu pop 
         Caption         =   "-"
      End
      Begin VB.Menu MenuEdito 
         Caption         =   "&Editor de Menu"
         Enabled         =   0   'False
         Shortcut        =   ^M
      End
      Begin VB.Menu menufsaf 
         Caption         =   "-"
      End
      Begin VB.Menu MenuLimpa 
         Caption         =   "Limpar Figura"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu MenuFormatar 
      Caption         =   "&Formatar"
      Begin VB.Menu MenuEnviarTraz 
         Caption         =   "&Enviar Objeto para Traz"
         Enabled         =   0   'False
         Shortcut        =   ^T
      End
      Begin VB.Menu ppp 
         Caption         =   "-"
      End
      Begin VB.Menu MenuEnviar 
         Caption         =   "&Enviar Objeto para Frente"
         Enabled         =   0   'False
         Shortcut        =   ^F
      End
   End
   Begin VB.Menu menuexecutar 
      Caption         =   "&Executar"
      Begin VB.Menu menuComplile 
         Caption         =   "&Executar          "
         Enabled         =   0   'False
         Shortcut        =   {F5}
      End
      Begin VB.Menu menfadfds 
         Caption         =   "-"
      End
      Begin VB.Menu menuopcoes 
         Caption         =   "&Instruções Iniciais"
         Enabled         =   0   'False
         Shortcut        =   ^O
      End
   End
End
Attribute VB_Name = "FrmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim A As String
Dim i As Long
Dim Oi As Long
Dim Row As String
Dim X As Long
Dim CboPass As Boolean
Dim Req As Boolean
Dim Passa As Boolean
Public NomeDoArquivoASerSalvo As String

Private Sub Ferramentas_CloseClick()
Dim H As Long, W As Long


W = FrmPrincipal.ScaleWidth + Ferramentas.Width
If FrmPropriedade.WindowState <> 1 Then
    FrmPropriedade.Left = W - FrmPropriedade.Width
End If
If FrmEx.WindowState <> 1 Then
    FrmEx.Left = W - FrmEx.Width
End If
Ferramentas.Visible = False
MenuFerra.Checked = False
End Sub

Private Sub MDIForm_Activate()
If Passa = True Then
    FrmPropriedade.Show
    FrmEx.Show
    Me.MousePointer = 0
    NoOpen = 0
    FrmApresentacao.Show 1
    Passa = False
    If NoOpen = 1 Then MenuAbrir_Click
    If NoOpen = 0 Then Exit Sub
    If NoOpen = 2 Then
        Abilita True
        Novo
    End If
End If
End Sub

Private Sub MDIForm_Load()
NomeDoArquivoASerSalvo = ""
Req = False
ReDim FrmTela(999) As New Form2
ContTela = 0
TelaAtual = -1
Req = True
Passa = True
MenuFerra.Checked = True
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
End
End Sub

Private Sub MenuAbrir_Click()
Abrir
End Sub

Private Sub MenuCod_Click()
FrmCodigo.Visible = True
FrmCodigo.TxtCod.SetFocus
End Sub

Private Sub menuComplile_Click()
If Dir("C:\TmpVar.Win") <> "" Then
    Kill "C:\TmpVar.Win"
End If
If ContTela = 0 Then
    MsgBox "Não a Telas para ser Executadas", vbInformation, App.Title
    Exit Sub
End If
Me.WindowState = 1
ContRun = 0
Start Nome_Da_Tela
End Sub

Private Sub MenuEdito_Click()
FrmMenu.Show 1
End Sub

Private Sub MenuEnviar_Click()
NovoObj.ZOrder vbBringToFront
Dim X As Long
For X = 0 To 7
    FrmTela(TelaAtual).Im(X).ZOrder vbSendToBack
    FrmTela(TelaAtual).Im(X).ZOrder vbBringToFront
Next X

End Sub

Private Sub MenuEnviarTraz_Click()
NovoObj.ZOrder vbSendToBack
Dim X As Long
For X = 0 To 7
    FrmTela(TelaAtual).Im(X).ZOrder vbSendToBack
Next X
End Sub

Private Sub MenuFerra_Click()
Ferramentas.Visible = True
FrmPropriedade.Height = (FrmPrincipal.ScaleHeight / 2) + 100
FrmPropriedade.Width = 3360
FrmPropriedade.Left = FrmPrincipal.ScaleWidth - FrmPropriedade.Width
FrmPropriedade.Top = (FrmPrincipal.ScaleHeight - FrmPropriedade.Height)
FrmEx.Height = (FrmPrincipal.ScaleHeight / 2) - 100
FrmEx.Width = 3360
FrmEx.Left = FrmPrincipal.ScaleWidth - FrmEx.Width
FrmEx.Top = 0
MenuFerra.Checked = True
End Sub

Private Sub MenuLimpa_Click()
NovoObj.Picture = LoadPicture("")
End Sub

Private Sub menuopcoes_Click()
FrmOpcoes.Show 1
'FrmCodigo.Visible = True
'If FrmCodigo.WindowState <> 1 Then
    'FrmCodigo.Height = 6000
    'FrmCodigo.Width = 6000
'End If
'FrmCodigo.Cbo.Caption = "Sistema"
'FrmCodigo.Env.Caption = "Inicialização"
'FrmCodigo.Cod.SetFocus
End Sub

Private Sub MenuProjeto_Click()
FrmEx.Visible = True
FrmEx.Prog.SetFocus
End Sub

Private Sub MenuProjetog_Click()
Me.MousePointer = 0
NoOpen = 0
FrmApresentacao.CmdAbrir.Visible = False
FrmApresentacao.Show 1
Passa = False
If NoOpen = 1 Then MenuAbrir_Click
If NoOpen = 0 Then Exit Sub
If NoOpen = 2 Then
    Abilita True
    Novo
End If
End Sub

Private Sub MenuPropriedade_Click()
FrmPropriedade.Visible = True
FrmPropriedade.Grid.SetFocus

End Sub

Private Sub MenuSair_Click()
End
End Sub

Private Sub MenuSalComo_Click()
SalProjetoMdb True
End Sub

Private Sub MenuSalvar_Click()
On Error Resume Next
SalProjetoMdb False
End Sub

Private Sub MenuTela_Click()
On Error Resume Next

A = ""

Inicio:

A = InputBox("Nome da Tela :", App.Title, A)
If A = "" Then
    Exit Sub
End If

If FrmEx.Prog.Nodes.Count = 0 Then GoTo Passa

Dim X1 As Long

For X1 = 0 To FrmEx.Prog.Nodes.Count - 1
    If UCase(A) = UCase(FrmEx.Prog.Nodes(X1 + 1).Text) Then
        MsgBox "Impossivel Criar uma tela com este nome, Pois ela já existe ! ! !", vbCritical, App.Title
        GoTo Inicio
        Exit Sub
    End If
Next X1

Passa:

FrmEx.Prog.Nodes.Add , , A, A, 1


With FrmCodigo.Eventos
    Dim nodX As Node
    
    .Nodes.Add "MOD", tvwChild, UCase(A), A, 1

    
    .Nodes.Add UCase(A), tvwChild, UCase(A + ".2"), "Ao Clicar 2 Vezes", 2
    .Nodes.Add UCase(A), tvwChild, UCase(A + ".1"), "Ao Clicar 1 Vezes", 2
    .Nodes.Add UCase(A), tvwChild, UCase(A + ".Ganhar"), "Ao Ganhar o focu", 2
    .Nodes.Add UCase(A), tvwChild, UCase(A + ".Perder"), "Ao Peder o focu", 2
    .Nodes.Add UCase(A), tvwChild, UCase(A + ".Red"), "Ao Redimecionar a tela", 2
    '.Nodes.Add UCase(A), tvwChild, UCase(A + ".Escrever"), "Ao Escrever", 2
    .Nodes.Add UCase(A), tvwChild, UCase(A + ".Fechar"), "Ao Fechar a Tela", 2
    .Nodes.Add UCase(A), tvwChild, UCase(A + ".Carregar"), "Ao Carregar a Tela", 2
    
End With

If ContTela = 0 Or ContTela = -1 Then
    Nome_Da_Tela = A
    ContTela = 0
End If

Max(ContTela) = True
Min(ContTela) = True
Fecha(ContTela) = True
FrmTela(ContTela).Width = 7305
FrmTela(ContTela).Nome.Caption = A
FrmTela(ContTela).Cont = ContTela
FrmTela(ContTela).Caption = "Tela " + Str(ContTela)
FrmTela(ContTela).Tag = A
FrmTela(ContTela).Show vbModeless
If ContTela = 0 Then
    TelaAtual = 0
End If
TelaAtual = ContTela
Set NovoObj = FrmTela(ContTela)

ContTela = ContTela + 1
End Sub

Private Sub T_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim X As Long
If TelaAtual = -1 Then
    Exit Sub
End If
For X = 1 To T.Buttons.Count
    T.Buttons(X).Value = tbrUnpressed
Next X
Button.Value = tbrPressed
If Button.Index = 1 Then
    FrmTela(TelaAtual).MousePointer = 0
    Exit Sub
End If
FrmTela(TelaAtual).MousePointer = 2
End Sub

Private Sub SalProjeto()
On Error Resume Next
Dim Arq As String, X As Byte, Diretorios As String, NameArq As String
Dim ArqAux As String, Xy As Long, Yx As Long, Ob As Object
Dim M As MenuItem, ArqCodigo As String

Com.FileName = ""
Com.Filter = "Projeto do Lego (*.Leg) |*.Leg|Todos os Arquivo (*.*)|*.*"
Com.ShowSave
If Com.FileName <> "" Then
    X = Len(Com.FileName)
    Do While X <> 0
        If Mid(Com.FileName, X, 1) = "\" Then
            Diretorios = Left(Com.FileName, X)
            Exit Do
        End If
        X = X - 1
    Loop
    NameArq = Right(Com.FileName, Len(Com.FileName) - X)
    ArqCodigo = Left(NameArq, Len(NameArq) - 4) + ".afs"
    
    If Dir(Com.FileName) <> "" Then
        If MsgBox("Arquivo Existente , Deseja Substituir ???", vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbYes Then
            Kill Com.FileName
'            Exit Sub
        End If
    End If
    Arq = Com.FileName
    Escreva "Inicialização", "Iniciar", Nome_Da_Tela, Arq
    Escreva "Inicialização", "Plataforma", Plataforma, Arq
    Escreva "Projeto", "Nome EXE", "Projeto.exe", Arq
    Escreva "Projeto", "Versão do Lego", "1.0", Arq
    Escreva "Projeto", "Nome do Projeto", NameArq, Arq
    Escreva "Projeto", "Titulo", Titulo_Pjt, Arq
    Escreva "Projeto", "Autor", Autor, Arq
    Escreva "Projeto", "Senha", CripSenha(Senha), Arq
    
    Escreva "Arquivos", "Dir", Diretorios, Arq
    
    For X = 0 To ContTela - 1
        Escreva "Arquivos", "Arq" + Format(Str(X), "000"), FrmTela(X).Tag + ".neo", Arq
        ArqAux = Diretorios + FrmTela(X).Tag + ".neo"
        Set Ob = FrmTela(X)
        Escreva Ob.Tag, "Nome", Ob.Tag, ArqAux
        Escreva Ob.Tag, "Tipo", Ob.Name, ArqAux
        Escreva Ob.Tag, "Texto", Ob.Text, ArqAux
        Escreva Ob.Tag, "Comprimir", Ob.Stretch, ArqAux
        Escreva Ob.Tag, "Figura", Ob.Picture, ArqAux
        Escreva Ob.Tag, "Borda", Ob.BorderStyle, ArqAux
        Escreva Ob.Tag, "Legenda", Ob.Caption, ArqAux
        Escreva Ob.Tag, "Tam-X", Ob.Height, ArqAux
        Escreva Ob.Tag, "Tam-Y", Ob.Width, ArqAux
        Escreva Ob.Tag, "Cor de Fundo", Ob.BackColor, ArqAux
        Escreva Ob.Tag, "Nome", Ob.Tag, ArqAux
        Escreva Ob.Tag, "Pox-X", Ob.Top, ArqAux
        Escreva Ob.Tag, "Pox-Y", Ob.Left, ArqAux
        Escreva Ob.Tag, "Tamanho", Ob.FontSize, ArqAux
        Escreva Ob.Tag, "Estilo", Ob.FontName, ArqAux
        Escreva Ob.Tag, "Cor de Letra", Ob.ForeColor, ArqAux
        Escreva Ob.Tag, "Local", Ob.ToolTipText, ArqAux
        Escreva Ob.Tag, "Ordem", Ob.TabIndex, ArqAux
        Escreva Ob.Tag, "3D", Ob.Appearance, ArqAux
        Escreva Ob.Tag, "Mascara", Ob.PasswordChar, ArqAux
        Escreva Ob.Tag, "MAX", Max(FrmTela(X).Cont.Caption), ArqAux
        Escreva Ob.Tag, "MIN", Min(FrmTela(X).Cont.Caption), ArqAux
        Escreva Ob.Tag, "Fechar", Fecha(FrmTela(X).Cont.Caption), ArqAux
        Xy = 0
        For Each Ob In FrmTela(X)
            If UCase(Ob.Name) = "XMENU" Or Ob.Name = "Im" Or Ob.Name = "Focus" Or Ob.Name = "S" Or Ob.Name = "Nome" Or Ob.Name = "Cont" Then
                GoTo proximo
            ElseIf Ob.Index = 0 Then
                GoTo proximo
            End If
            Escreva "Objetos", "Nome" + Trim(Str(Xy)), Ob.Tag, ArqAux
            Escreva Ob.Tag, "Tipo", Ob.Name, ArqAux
            Escreva Ob.Tag, "Texto", Ob.Text, ArqAux
            Escreva Ob.Tag, "Comprimir", Ob.Stretch, ArqAux
            Escreva Ob.Tag, "Figura", Ob.ToolTipText, ArqAux
            Escreva Ob.Tag, "Borda", Ob.BorderStyle, ArqAux
            Escreva Ob.Tag, "Legenda", Ob.Caption, ArqAux
            Escreva Ob.Tag, "Tam-X", Ob.Height, ArqAux
            Escreva Ob.Tag, "Tam-Y", Ob.Width, ArqAux
            Escreva Ob.Tag, "Cor de Fundo", Ob.BackColor, ArqAux
            Escreva Ob.Tag, "Nome", Ob.Tag, ArqAux
            Escreva Ob.Tag, "Pox-X", Ob.Top, ArqAux
            Escreva Ob.Tag, "Pox-Y", Ob.Left, ArqAux
            Escreva Ob.Tag, "Tamanho", Ob.FontSize, ArqAux
            Escreva Ob.Tag, "Estilo", Ob.FontName, ArqAux
            Escreva Ob.Tag, "Cor de Letra", Ob.ForeColor, ArqAux
            Escreva Ob.Tag, "Local", Ob.ToolTipText, ArqAux
            Escreva Ob.Tag, "Ordem", Ob.TabIndex, ArqAux
            Escreva Ob.Tag, "3D", Ob.Appearance, ArqAux
            Escreva Ob.Tag, "Mascara", Ob.PasswordChar, ArqAux
            Xy = Xy + 1

proximo:
                If UCase(Ob.Name) = "XMENU" Then
                    Escreva "Menu", "Fonte", FrmTela(X).Xmenu.ItemsFont, ArqAux
                    Escreva "Menu", "Borda", FrmTela(X).Xmenu.Style, ArqAux
                    Escreva "Menu", "Selecao", FrmTela(X).Xmenu.HighLightStyle, ArqAux
                    
                    For Yx = 0 To FrmTela(X).Xmenu.MenuTree.Count
                        Set M = FrmTela(X).Xmenu.MenuTree(Yx)
                        Escreva "Menu", "Legenda" + Trim(Str(Yx)), M.Caption, ArqAux
                        Escreva "Menu", "Ident" + Trim(Str(Yx)), M.Ident, ArqAux
                        Escreva "Menu", "Chave" + Trim(Str(Yx)), M.Name, ArqAux
                        Escreva "Menu", "Root" + Trim(Str(Yx)), M.RootIndex, ArqAux
                    Next Yx
                End If
            Next

    Next X
    SalvarEventos1 Diretorios + ArqCodigo
End If
End Sub

Private Sub Abrir1()
On Error Resume Next
Dim Arq As String, X As Byte, Diretorios As String, NameArq As String
Dim ArqAux As String, Xy As Long, Yx As Long, Ob As Object
Dim M As MenuItem, ArqCodigo As String, Nometela As String
Dim Tipo As String, TipoX As String
Dim Hp As Long, A As String, Index As Long
Dim SenhaAux As String
Dim A1 As String, A2 As String

Com.FileName = ""
Com.Filter = "Projeto do Lego (*.Leg) |*.Leg|Todos os Arquivo (*.*)|*.*"
Com.ShowOpen

If Com.FileName <> "" Then
    X = Len(Com.FileName)
    Do While X <> 0
        If Mid(Com.FileName, X, 1) = "\" Then
            Diretorios = Left(Com.FileName, X)
            Exit Do
        End If
        X = X - 1
    Loop
    NameArq = Right(Com.FileName, Len(Com.FileName) - X)
    ArqCodigo = Left(NameArq, Len(NameArq) - 4) + ".afs"
    
    If Dir(Com.FileName) = "" Then
        MsgBox "Arquivo Invalido", vbQuestion
        Exit Sub
    End If
    Arq = Com.FileName
    Nome_Da_Tela = Ler("Inicialização", "Iniciar", "", Arq)
    Plataforma = Ler("Inicialização", "Plataforma", "", Arq)
    'Escreva "Projeto", "Nome EXE", "Projeto.exe", Arq
    'Escreva "Projeto", "Versão do Lego", "1.0", Arq
    'Escreva "Projeto", "Nome do Projeto", NameArq, Arq
    SenhaAux = Senha
    A1 = Ler("Projeto", "Titulo", "", Arq)
    A2 = Ler("Projeto", "Autor", "", Arq)
    Senha = DescpSenha(Ler("Projeto", "Senha", "", Arq))
    If Senha <> "" Then
        SenhaOk = False
        FrmSenha.Auto.Caption = "Autor : " + A2
        FrmSenha.Caption = "Senha : " + A1
        FrmSenha.Show 1
        If SenhaOk = False Then
            Senha = SenhaAux
            Exit Sub
        End If
    End If
    Titulo_Pjt = Ler("Projeto", "Titulo", "", Arq)
    Autor = Ler("Projeto", "Autor", "", Arq)
   
    Diretorios = Ler("Arquivos", "Dir", App.Path, Arq)
    X = 0
    ContTela = 0
    X = 0
    FrmCodigo.Eventos.Refresh
    Do While FrmCodigo.Eventos.Nodes.Count <> 0
        FrmCodigo.Eventos.Nodes.Remove 1
        X = X + 1
    Loop
    FrmCodigo.Eventos.Nodes.Add , , "MOD", "Modulos", 4
    X = 0
    FechaGeral = False
    Do While FrmEx.Prog.Nodes.Count <> 0
        Unload FrmTela(X)
        FrmEx.Prog.Nodes.Remove 1
        X = X + 1
    Loop
    FechaGeral = True
    X = 0
    Do While True
    
proximo:
        Nometela = Ler("Arquivos", "Arq" + Format(Str(X), "000"), ";=;", Arq)
        If Nometela = ";=;" Then
            Exit Do
        End If
            
        ArqAux = Diretorios + Nometela
        If Dir(ArqAux) = "" Then
            MsgBox "Impossivel Localizar M Arquivo " + ArqAux, vbCritical, App.Title
            X = X + 1
            GoTo proximo
        End If
        Err.Number = 0
        Load FrmTela(X)
        FrmTela(X).Tag = Left(Nometela, Len(Nometela) - 4)
        FrmTela(X).Visible = False
        FrmTela(X).Cont.Caption = X
        Set Ob = FrmTela(X)
        Ob.Tag = Ler(Ob.Tag, "Nome", "", ArqAux)
        Ob.BorderStyle = Ler(Ob.Tag, "Borda", "", ArqAux)
        Ob.Caption = Ler(Ob.Tag, "Legenda", Ob.Caption, ArqAux)
        Ob.Height = Ler(Ob.Tag, "Tam-X", Ob.Height, ArqAux)
        Ob.Width = Ler(Ob.Tag, "Tam-Y", Ob.Width, ArqAux)
        Ob.BackColor = Ler(Ob.Tag, "Cor de Fundo", Ob.BackColor, ArqAux)
        Ob.Top = Ler(Ob.Tag, "Pox-X", Ob.Top, ArqAux)
        Ob.Left = Ler(Ob.Tag, "Pox-Y", Ob.Left, ArqAux)
        Ob.FontSize = Ler(Ob.Tag, "Tamanho", Ob.FontSize, ArqAux)
        Ob.FontName = Ler(Ob.Tag, "Estilo", Ob.FontName, ArqAux)
        Ob.ForeColor = Ler(Ob.Tag, "Cor de Letra", Ob.ForeColor, ArqAux)
        Ob.Appearance = Ler(Ob.Tag, "3D", "1", ArqAux)
        Ob.Visible = True
        Max(X) = Ler(Ob.Tag, "MAX", "FALSE", ArqAux)
        Min(X) = Ler(Ob.Tag, "MIN", "FALSE", ArqAux)
        Fecha(X) = Ler(Ob.Tag, "Fechar", "FALSE", ArqAux)
        
        A = Ob.Tag
        FrmEx.Prog.Nodes.Add , , A, A, 1


        With FrmCodigo.Eventos
            Dim nodX As Node
            
            .Nodes.Add "MOD", tvwChild, UCase(A), A, 1
        
            
            .Nodes.Add UCase(A), tvwChild, UCase(A + ".2"), "Ao Clicar 2 Vezes", 2
            .Nodes.Add UCase(A), tvwChild, UCase(A + ".1"), "Ao Clicar 1 Vezes", 2
            .Nodes.Add UCase(A), tvwChild, UCase(A + ".Ganhar"), "Ao Ganhar M focu", 2
            .Nodes.Add UCase(A), tvwChild, UCase(A + ".Perder"), "Ao Peder M focu", 2
            .Nodes.Add UCase(A), tvwChild, UCase(A + ".Red"), "Ao Redimecionar a tela", 2
            .Nodes.Add UCase(A), tvwChild, UCase(A + ".Escrever"), "Ao Escrever", 2
            .Nodes.Add UCase(A), tvwChild, UCase(A + ".Fechar"), "Ao Fechar a Tela", 2
            .Nodes.Add UCase(A), tvwChild, UCase(A + ".Carregar"), "Ao Carregar a Tela", 2
            
        End With
        ContTela = ContTela + 1

        Xy = 0
        For Hp = 0 To 9999
            TipoX = Ler("Objetos", "Nome" + Trim(Str(Hp)), ";=;", ArqAux)
            If TipoX = ";=;" Then
                Exit For
            End If
            Index = -1
            Tipo = Ler(TipoX, "Tipo", "", ArqAux)
            If Tipo = "" Then Exit For
   
            If Tipo = "Cmd" Then
                Index = 0
                Load FrmTela(X).Cmd(FrmTela(X).Cmd.Count)
                Set Ob = FrmTela(X).Cmd(FrmTela(X).Cmd.Count - 1)
            ElseIf Tipo = "Fm" Then
                Index = 1
                Load FrmTela(X).Fm(FrmTela(X).Fm.Count)
                Set Ob = FrmTela(X).Fm(FrmTela(X).Fm.Count - 1)
            ElseIf Tipo = "Img" Then
                Index = 2
                Load FrmTela(X).Img(FrmTela(X).Img.Count)
                Set Ob = FrmTela(X).Img(FrmTela(X).Img.Count - 1)
                Ob.Picture = Ob.Picture
            ElseIf Tipo = "Lbl" Then
                Index = 3
                Load FrmTela(X).Lbl(FrmTela(X).Lbl.Count)
                Set Ob = FrmTela(X).Lbl(FrmTela(X).Lbl.Count - 1)
            ElseIf Tipo = "Chk" Then
                Index = 4
                Load FrmTela(X).Chk(FrmTela(X).Chk.Count)
                Set Ob = FrmTela(X).Chk(FrmTela(X).Chk.Count - 1)
            ElseIf Tipo = "Cbo" Then
                Index = 5
                Load FrmTela(X).Cbo(FrmTela(X).Cbo.Count)
                Set Ob = FrmTela(X).Cbo(FrmTela(X).Cbo.Count - 1)
            ElseIf Tipo = "Txt" Then
                Index = 6
                Load FrmTela(X).Txt(FrmTela(X).Txt.Count)
                Set Ob = FrmTela(X).Txt(FrmTela(X).Txt.Count - 1)
            ElseIf Tipo = "Lst" Then
                Index = 7
                Load FrmTela(X).Lst(FrmTela(X).Lst.Count)
                Set Ob = FrmTela(X).Lst(FrmTela(X).Lst.Count - 1)
            End If
            
                
            Ob.Tag = TipoX
            Ob.Text = Ler(Ob.Tag, "Texto", "", ArqAux)
            Ob.Stretch = Ler(Ob.Tag, "Comprimir", Ob.Stretch, ArqAux)
            Ob.Picture = LoadPicture(Ler(Ob.Tag, "Figura", Ob.Picture, ArqAux))
            Ob.BorderStyle = Ler(Ob.Tag, "Borda", Ob.BorderStyle, ArqAux)
            Ob.Caption = Ler(Ob.Tag, "Legenda", Ob.Caption, ArqAux)
            Ob.Height = Ler(Ob.Tag, "Tam-X", Ob.Height, ArqAux)
            Ob.Width = Ler(Ob.Tag, "Tam-Y", Ob.Width, ArqAux)
            Ob.BackColor = Ler(Ob.Tag, "Cor de Fundo", Ob.BackColor, ArqAux)
            Ob.Top = Ler(Ob.Tag, "Pox-X", Ob.Top, ArqAux)
            Ob.Left = Ler(Ob.Tag, "Pox-Y", Ob.Left, ArqAux)
            Ob.FontSize = Ler(Ob.Tag, "Tamanho", Ob.FontSize, ArqAux)
            Ob.FontName = Ler(Ob.Tag, "Estilo", Ob.FontName, ArqAux)
            Ob.ForeColor = Ler(Ob.Tag, "Cor de Letra", Ob.ForeColor, ArqAux)
            Ob.ToolTipText = Ler(Ob.Tag, "Local", Ob.ToolTipText, ArqAux)
            Ob.TabIndex = Ler(Ob.Tag, "Ordem", Ob.TabIndex, ArqAux)
            Ob.Appearance = Ler(Ob.Tag, "3D", Ob.Appearance, ArqAux)
            Ob.PasswordChar = Ler(Ob.Tag, "Mascara", Ob.PasswordChar, ArqAux)
            Ob.Visible = True
            FrmCodigo.Eventos.Nodes.Add UCase(FrmTela(X).Tag), tvwChild, UCase(FrmTela(X).Tag + "." + Ob.Tag), Ob.Tag, 3

            With FrmCodigo.Eventos
            
                .Nodes.Add UCase(FrmTela(X).Tag + "." + Ob.Tag), tvwChild, UCase(FrmTela(X).Tag + "." + Ob.Tag + ".1"), "Ao Clicar 1 Vezes", 2
                If Index = 7 Or Index = 6 Or Index = 5 Or Index = 3 Or Index = 2 Or Index = 1 Then
                    .Nodes.Add UCase(FrmTela(X).Tag + "." + Ob.Tag), tvwChild, UCase(FrmTela(X).Tag + "." + Ob.Tag + ".2"), "Ao Clicar 2 Vezes", 2
                End If
                .Nodes.Add UCase(FrmTela(X).Tag + "." + Ob.Tag), tvwChild, UCase(FrmTela(X).Tag + "." + Ob.Tag + ".Ganhar"), "Ao Ganhar M focu", 2
                .Nodes.Add UCase(FrmTela(X).Tag + "." + Ob.Tag), tvwChild, UCase(FrmTela(X).Tag + "." + Ob.Tag + ".Perder"), "Ao Perder M focu", 2
                If Index = 6 Or 3 Then
                    .Nodes.Add UCase(FrmTela(X).Tag + "." + Ob.Tag), tvwChild, UCase(FrmTela(X).Tag + "." + Ob.Tag + ".Escrever"), "Ao Escrever", 2
                End If
                
            End With
            
            Xy = Xy + 1
        Next Hp
        FrmTela(X).Xmenu.ItemsFont = Ler("Menu", "Fonte", FrmTela(X).Xmenu.ItemsFont, ArqAux)
        FrmTela(X).Xmenu.Style = Ler("Menu", "Borda", FrmTela(X).Xmenu.Style, ArqAux)
        FrmTela(X).Xmenu.HighLightStyle = Ler("Menu", "Selecao", FrmTela(X).Xmenu.HighLightStyle, ArqAux)
        FrmTela(X).Visible = True
        FrmTela(X).Show
        FrmTela(X).Refresh
        For Yx = 1 To 9999
            Tipo = Ler("Menu", "Legenda" + Trim(Str(Yx)), ";=;", ArqAux)
            If Tipo = ";=;" Then Exit For
            Set M = New MenuItem
            M.Caption = Tipo
            M.Ident = Ler("Menu", "Ident" + Trim(Str(Yx)), M.Ident, ArqAux)
            M.Name = Ler("Menu", "Chave" + Trim(Str(Yx)), M.Name, ArqAux)
            M.RootIndex = Ler("Menu", "Root" + Trim(Str(Yx)), M.RootIndex, ArqAux)
            Nome = M.Caption
            InX = M.Ident
            M.Ident = InX
            M.Caption = Nome
            M.Name = Nome
            M.Accelerator = Nome
            M.Description = Nome
            If M.Ident = 0 Then
                FrmCodigo.Eventos.Nodes.Add UCase(FrmTela(TelaAtual).Tag), tvwChild, UCase(FrmTela(TelaAtual).Tag + "." + M.Caption + Trim(Str(M.Ident))), M.Caption, 5
            ElseIf M.Ident = 1 Then
                FrmCodigo.Eventos.Nodes.Add NomeMenu1, tvwChild, UCase(FrmTela(TelaAtual).Tag + "." + M.Caption + Trim(Str(M.Ident))), M.Caption, 5
            ElseIf M.Ident = 2 Then
                FrmCodigo.Eventos.Nodes.Add NomeMenu2, tvwChild, UCase(FrmTela(TelaAtual).Tag + "." + M.Caption + Trim(Str(M.Ident))), M.Caption, 5
            ElseIf M.Ident = 3 Then
                FrmCodigo.Eventos.Nodes.Add NomeMenu3, tvwChild, UCase(FrmTela(TelaAtual).Tag + "." + M.Caption + Trim(Str(M.Ident))), M.Caption, 5
            ElseIf M.Ident = 4 Then
                FrmCodigo.Eventos.Nodes.Add NomeMenu4, tvwChild, UCase(FrmTela(TelaAtual).Tag + "." + M.Caption + Trim(Str(M.Ident))), M.Caption, 5
            ElseIf M.Ident = 5 Then
                FrmCodigo.Eventos.Nodes.Add NomeMenu5, tvwChild, UCase(FrmTela(TelaAtual).Tag + "." + M.Caption + Trim(Str(M.Ident))), M.Caption, 5
            ElseIf M.Ident = 6 Then
                FrmCodigo.Eventos.Nodes.Add NomeMenu6, tvwChild, UCase(FrmTela(TelaAtual).Tag + "." + M.Caption + Trim(Str(M.Ident))), M.Caption, 5
            End If
            If M.Ident = 0 Then
                NomeMenu1 = UCase(FrmTela(TelaAtual).Tag + "." + M.Caption + Trim(Str(M.Ident)))
            ElseIf M.Ident = 1 Then
                NomeMenu2 = UCase(FrmTela(TelaAtual).Tag + "." + M.Name + Trim(Str(M.Ident)))
            ElseIf M.Ident = 2 Then
                NomeMenu3 = UCase(FrmTela(TelaAtual).Tag + "." + M.Name + Trim(Str(M.Ident)))
            ElseIf M.Ident = 3 Then
                NomeMenu4 = UCase(FrmTela(TelaAtual).Tag + "." + M.Name + Trim(Str(M.Ident)))
            ElseIf M.Ident = 4 Then
                NomeMenu5 = UCase(FrmTela(TelaAtual).Tag + "." + M.Name + Trim(Str(M.Ident)))
            ElseIf M.Ident = 5 Then
                NomeMenu6 = UCase(FrmTela(TelaAtual).Tag + "." + M.Name + Trim(Str(M.Ident)))
            End If
            FrmTela(X).Xmenu.MenuTree.Add M
        Next Yx
        FrmTela(X).Xmenu.Refresh
        FrmTela(X).SetFocus
        X = X + 1
    Loop
    BuscaEventosArq Diretorios + ArqCodigo
    Abilita True
End If

End Sub

Private Function CripSenha(Senha As String)
On Error Resume Next
Dim NovaSenha As String
Dim X As Long

NovaSenha = ""
If Senha = "" Then
    CripSenha = ""
    Exit Function
End If
For X = 1 To Len(Senha) + 1
    NovaSenha = NovaSenha + Chr(Asc(Mid(Senha, X, 1)) + 37)
Next X
CripSenha = NovaSenha
End Function
Private Function DescpSenha(Senha As String)
On Error Resume Next
Dim NovaSenha As String
Dim X As Long

NovaSenha = ""
If Senha = "" Then
    DescpSenha = ""
    Exit Function
End If
For X = 1 To Len(Senha)
    NovaSenha = NovaSenha + Chr(Asc(Mid(Senha, X, 1)) - 37)
Next X
DescpSenha = NovaSenha
End Function

Private Sub SalProjetoMdb(Tipo As Boolean)
On Error Resume Next
Dim Arq As String, X As Byte, Diretorios As String, NameArq As String
Dim ArqAux As String, Xy As Long, Yx As Long, Ob As Object
Dim M As MenuItem, ArqCodigo As String
Dim Rs As Recordset
Inicio:
Com.FileName = ""
Com.Filter = "Projeto do Lego (*.Afs) |*.afs|Todos os Arquivo (*.*)|*.*"
If Tipo = True Then
    Com.ShowSave
Else
    If NomeDoArquivoASerSalvo = "" Then
        Com.ShowOpen
        Tipo = True
    Else
        Com.FileName = NomeDoArquivoASerSalvo
    End If
End If
If Com.FileName <> "" Then
    'X = Len(Com.FileName)
'    Do While X <> 0
   '     If Mid(Com.FileName, X, 1) = "\" Then
   '         Diretorios = Left(Com.FileName, X)
   '         Exit Do
   '     End If
    '    X = X - 1
    'Loop
    NameArq = Com.FileName
    
    If Dir(Com.FileName) <> "" Then
        If Tipo = True Then
            If MsgBox("Arquivo Existente , Deseja Substituir ???", vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbYes Then
                Kill Com.FileName
    '            Exit Sub
            Else
                GoTo Inicio
            End If
        Else
            Kill Com.FileName
        End If
    End If
    Screen.MousePointer = 11
    Arq = Com.FileName
    CriarBanco Arq
    Set Rs = Banco.OpenRecordset("Config")
    
    Rs.AddNew
    Rs!Banco = LocalBancodeDados
    Rs!Iniciar = Nome_Da_Tela
    Rs!Plataforma = Plataforma
    Rs!EXE = "Projeto.exe"
    Rs!Titulo = Titulo_Pjt
    Rs!Autor = Autor
    Rs!Senha = CripSenha(Senha)
    Rs!Esc = TeclaEsc
    Rs!Enter = TeclaEnter
    Rs.Update
    
    For X = 0 To ContTela - 1
        Set Ob = FrmTela(X)
        CriarEstrutura FrmTela(X).Tag
        Set Rs = Banco.OpenRecordset(FrmTela(X).Tag)
        Rs.AddNew
        Rs!Nome = Ob.Tag
        Rs!Tipo = Ob.Name
        Rs!Borda = Ob.BorderStyle
        Rs!Legenda = Ob.Caption
        Rs.Fields("TamX") = Ob.Height
        Rs.Fields("TamY") = Ob.Width
        Rs.Fields("Cor Fundo") = Ob.BackColor
        Rs.Fields("PoxX") = Ob.Top
        Rs.Fields("PoxY") = Ob.Left
        Rs.Fields("Tamanho") = Val(Ob.FontSize)
        Rs.Fields("Estilo") = Ob.FontName
        Rs.Fields("Cor da Letra") = Ob.ForeColor
        Rs.Fields("3D") = Ob.Appearance
        Rs.Fields("BotaoMAX") = Max(FrmTela(X).Cont.Caption)
        Rs.Fields("BotaoMIN") = Min(FrmTela(X).Cont.Caption)
        Rs.Fields("BotaoFechar") = Fecha(FrmTela(X).Cont.Caption)
        Rs!Imagem = BuscaImg(Ob.Icon)
        Rs.Update
        Xy = 0
        For Each Ob In FrmTela(X)
            If UCase(Ob.Name) = "XMENU" Or Ob.Name = "Im" Or Ob.Name = "Focus" Or Ob.Name = "S" Or Ob.Name = "Nome" Or Ob.Name = "Cont" Then
                GoTo proximo
            ElseIf Ob.Index = 0 Then
                GoTo proximo
            End If
            Set Rs = Banco.OpenRecordset(FrmTela(X).Tag)
            Rs.AddNew
            Rs!Nome = Ob.Tag
            Rs!Tipo = Ob.Name
            Rs!Texto = Ob.Text
            Rs!Comprimir = Ob.Stretch
            Rs!Imagem = BuscaImg(Ob.Picture)
            Rs!Borda = Ob.BorderStyle
            Rs!Legenda = Ob.Caption
            Rs.Fields("TamX") = Ob.Height
            Rs.Fields("TamY") = Ob.Width
            Rs.Fields("Cor Fundo") = Ob.BackColor
            Rs.Fields("PoxX") = Ob.Top
            Rs.Fields("PoxY") = Ob.Left
            Rs.Fields("Tamanho") = Ob.FontSize
            Rs.Fields("Estilo") = Ob.FontName
            Rs.Fields("Cor dd Letra") = Ob.ForeColor
            Rs.Fields("3D") = Ob.Appearance
            Rs.Fields("Ordem") = Ob.TabIndex
            Rs.Fields("Mascara") = Ob.PasswordChar
            Rs.Fields("Order") = Ob.TabIndex
            Rs.Update
            Xy = Xy + 1

proximo:
                If UCase(Ob.Name) = "XMENU" Then
                    Set Rs = Banco.OpenRecordset("Menu-" + FrmTela(X).Tag)
                    Rs.AddNew
                    Rs!Legenda = FrmTela(X).Xmenu.ItemsFont
                    Rs!Chave = FrmTela(X).Xmenu.Style
                    Rs!Root = FrmTela(X).Xmenu.HighLightStyle
                    Rs.Update
                    For Yx = 1 To FrmTela(X).Xmenu.MenuTree.Count
                        Set M = FrmTela(X).Xmenu.MenuTree(Yx)
                        Rs.AddNew
                        Rs!Legenda = M.Caption
                        Rs!Ident = M.Ident
                        Rs!Chave = M.Name
                        Rs!Root = M.RootIndex
                        Rs.Update
                    Next Yx
                End If
            Next

    Next X
    SalvarEventos
    Banco.Close
    Screen.MousePointer = 0
    NomeDoArquivoASerSalvo = Com.FileName
End If
Screen.MousePointer = 0
End Sub


Private Function Abilita(Tipo As Boolean)
On Error Resume Next
Dim X As Long
MenuSalvar.Enabled = Tipo
MenuSalComo.Enabled = Tipo
MenuCompli.Enabled = Tipo
MenuProjeto.Enabled = Tipo
MenuPropriedade.Enabled = Tipo
MenuCod.Enabled = Tipo
MenuFerra.Enabled = Tipo
MenuEdito.Enabled = Tipo
MenuEnviarTraz.Enabled = Tipo
MenuEnviar.Enabled = Tipo
menuComplile.Enabled = Tipo
menuopcoes.Enabled = Tipo
'MenuBancoDados.Enabled = Tipo
MenuTela.Enabled = Tipo
T.Enabled = Tipo
Toolbar1.Buttons(3).Enabled = Tipo
Toolbar1.Buttons(7).Enabled = Tipo
Toolbar1.Buttons(8).Enabled = Tipo
Toolbar1.Buttons(1).ButtonMenus(2).Enabled = Tipo
'Toolbar1.Buttons(1).ButtonMenus(3).Enabled = Tipo
'Toolbar1.Buttons(1).ButtonMenus(4).Enabled = Tipo
MenuLimpa.Enabled = Tipo

End Function

Private Sub Novo()
Dim X As Long
X = 1
FechaGeral = False
Do While FrmCodigo.Eventos.Nodes.Count <> 0
    FrmCodigo.Eventos.Nodes.Remove X
    X = X + 1
Loop
X = 1
Do While FrmEx.Prog.Nodes.Count <> 0
    Unload FrmTela(X - 1)
    FrmEx.Prog.Nodes.Remove 1
    X = X + 1
Loop

TeclaEnter = 0
TeclaEsc = 0
FechaGeral = True
ContTela = -1
TelaAtual = -1
FrmCodigo.Eventos.Nodes.Add , , "MOD", "Modulos", 4
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
If Button.Index = 1 Then
    If Toolbar1.Buttons(1).ButtonMenus(2).Enabled = True Then
        Call MenuProjetog_Click
    Else
        Call MenuTela_Click
        
    End If
ElseIf Button.Index = 2 Then
    Abrir
ElseIf Button.Index = 3 Then
    SalProjetoMdb False
ElseIf Button.Index = 7 Then
    menuComplile_Click
End If
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
If ButtonMenu.Index = 1 Then
    Call MenuProjetog_Click
ElseIf ButtonMenu.Index = 2 Then
    Call MenuTela_Click
End If
End Sub
Private Sub Abrir()
On Error Resume Next
Dim Arq As String, X As Byte, Diretorios As String, NameArq As String
Dim ArqAux As String, Xy As Long, Yx As Long, Ob As Object
Dim M As MenuItem, ArqCodigo As String, Nometela As String
Dim Tipo As String, TipoX As String
Dim Hp As Long, A As String, Index As Long
Dim SenhaAux As String
Dim A1 As String, A2 As String
Dim Rs As Recordset
Dim TabelaIndex As Long

Com.FileName = ""
Com.Filter = "Projeto do Lego (*.Afs) |*.Afs|Todos os Arquivo (*.*)|*.*"
Com.ShowOpen

If Com.FileName <> "" Then
    X = Len(Com.FileName)
    'Do While X <> 0
    '    If Mid(Com.FileName, X, 1) = "\" Then
    '        Diretorios = Left(Com.FileName, X)
    '        Exit Do
    '    End If
    '    X = X - 1
    'Loop
    NameArq = Com.FileName
    NomeDoArquivoASerSalvo = Com.FileName
    If Dir(Com.FileName) = "" Then
        MsgBox "Arquivo Invalido", vbQuestion
        Exit Sub
    End If
    Arq = Com.FileName
    Err.Number = 0
    
    'Abre o Arquivo
    Set Banco = OpenDatabase(Arq)
    If Err.Number <> 0 Then
        MsgBox "Arquivo em Formato Invalido ! ! !", vbCritical, App.Title
        Banco.Close
        Exit Sub
    End If
    
    'Abre as Configurações
    Set Rs = Banco.OpenRecordset("Config")
    If Rs.RecordCount = 0 Then
        MsgBox "Impossivel localizar Paramentos a do Sistema", vbCritical, App.Title
        Rs.Close
        Banco.Close
        Exit Sub
    End If
    
    SenhaAux = Senha
    A1 = Rs!Titulo
    A2 = Rs!Autor
    Senha = DescpSenha(Rs!Senha)
    If Senha <> "" Then
        SenhaOk = False
        FrmSenha.Auto.Caption = "Autor : " + A2
        FrmSenha.Caption = "Senha : " + A1
        FrmSenha.Show 1
        If SenhaOk = False Then
            Senha = SenhaAux
            Exit Sub
        End If
    End If
    Nome_Da_Tela = Rs!Iniciar
    Plataforma = Rs!Plataforma
    Titulo_Pjt = Rs!Titulo
    Autor = Rs!Autor
    TeclaEnter = Rs!Enter
    TeclaEsc = Rs!Esc
    LocalBancodeDados = Rs!Banco
    X = 0
    ContTela = 0
    X = 0
    FrmCodigo.Eventos.Refresh
    Do While FrmCodigo.Eventos.Nodes.Count <> 0
        FrmCodigo.Eventos.Nodes.Remove 1
        X = X + 1
    Loop
    FrmCodigo.Eventos.Nodes.Add , , "MOD", "Modulos", 4
    X = 0
    FechaGeral = False
    Do While FrmEx.Prog.Nodes.Count <> 0
        Unload FrmTela(X)
        FrmEx.Prog.Nodes.Remove 1
        X = X + 1
    Loop
    FechaGeral = True
    X = 0
    TabelaIndex = 0
    Do While TabelaIndex <> Banco.TableDefs.Count
proximo:

        Nometela = Banco.TableDefs(TabelaIndex).Name
        
        If Nometela = "Config" Or Nometela = "Lego" Or Nometela = "MSysACEs" Or Nometela = "MSysModules" Or Nometela = "MSysModules2" Or Nometela = "MSysObjects" Or Nometela = "MSysQueries" Or Nometela = "MSysRelationships" Or Left(Nometela, 5) = "Menu-" Then
            TabelaIndex = TabelaIndex + 1
            If TabelaIndex >= Banco.TableDefs.Count Then Exit Do
            GoTo proximo
        End If
            
        'Inicio da Criação
        
        Err.Number = 0
        Load FrmTela(X)
        FrmTela(X).Tag = Nometela
        FrmTela(X).Visible = False
        FrmTela(X).Cont.Caption = X
        Set Ob = FrmTela(X)
        Set Rs = Banco.OpenRecordset(Nometela)
        Ob.Appearance = Rs.Fields("3D")
        Ob.Tag = Nometela
        Ob.BorderStyle = Rs!Borda
        Ob.Caption = Rs!Legenda
        Ob.Height = Rs.Fields("TamX")
        Ob.Width = Rs.Fields("TamY")
        Ob.BackColor = Rs.Fields("Cor Fundo")
        Ob.Top = Rs.Fields("PoxX")
        Ob.Left = Rs.Fields("PoxY")
        Ob.FontSize = Rs.Fields("Tamanho")
        Ob.FontName = Rs.Fields("Estilo")
        Ob.ForeColor = Rs.Fields("Cor da Letra")
        AbreFig Rs!Imagem, Ob, False
        'Ob.Visible = True
        Max(X) = Rs.Fields("BotaoMAX")
        Min(X) = Rs.Fields("BotaoMIN")
        Fecha(X) = Rs.Fields("BotaoFechar")
        
        A = Ob.Tag
        FrmEx.Prog.Nodes.Add , , A, A, 1


        With FrmCodigo.Eventos
            Dim nodX As Node
            
            .Nodes.Add "MOD", tvwChild, UCase(A), A, 1
        
            
            .Nodes.Add UCase(A), tvwChild, UCase(A + ".2"), "Ao Clicar 2 Vezes", 2
            .Nodes.Add UCase(A), tvwChild, UCase(A + ".1"), "Ao Clicar 1 Vezes", 2
            .Nodes.Add UCase(A), tvwChild, UCase(A + ".Ganhar"), "Ao Ganhar M focu", 2
            .Nodes.Add UCase(A), tvwChild, UCase(A + ".Perder"), "Ao Peder M focu", 2
            .Nodes.Add UCase(A), tvwChild, UCase(A + ".Red"), "Ao Redimecionar a tela", 2
            .Nodes.Add UCase(A), tvwChild, UCase(A + ".Escrever"), "Ao Escrever", 2
            .Nodes.Add UCase(A), tvwChild, UCase(A + ".Fechar"), "Ao Fechar a Tela", 2
            .Nodes.Add UCase(A), tvwChild, UCase(A + ".Carregar"), "Ao Carregar a Tela", 2
            
        End With
        ContTela = ContTela + 1

        Xy = 0
        Rs.MoveNext
        If Rs.EOF = True Then GoTo ProximaTela
        
        Do While Not Rs.EOF
            Index = -1
            Tipo = Rs!Tipo
  
            If Tipo = "Cmd" Then
                Index = 0
                Load FrmTela(X).Cmd(FrmTela(X).Cmd.Count)
                Set Ob = FrmTela(X).Cmd(FrmTela(X).Cmd.Count - 1)
            ElseIf Tipo = "Fm" Then
                Index = 1
                Load FrmTela(X).Fm(FrmTela(X).Fm.Count)
                Set Ob = FrmTela(X).Fm(FrmTela(X).Fm.Count - 1)
            ElseIf Tipo = "Img" Then
                Index = 2
                Load FrmTela(X).Img(FrmTela(X).Img.Count)
                Set Ob = FrmTela(X).Img(FrmTela(X).Img.Count - 1)
            ElseIf Tipo = "Lbl" Then
                Index = 3
                Load FrmTela(X).Lbl(FrmTela(X).Lbl.Count)
                Set Ob = FrmTela(X).Lbl(FrmTela(X).Lbl.Count - 1)
            ElseIf Tipo = "Chk" Then
                Index = 4
                Load FrmTela(X).Chk(FrmTela(X).Chk.Count)
                Set Ob = FrmTela(X).Chk(FrmTela(X).Chk.Count - 1)
            ElseIf Tipo = "Cbo" Then
                Index = 5
                Load FrmTela(X).Cbo(FrmTela(X).Cbo.Count)
                Set Ob = FrmTela(X).Cbo(FrmTela(X).Cbo.Count - 1)
            ElseIf Tipo = "Txt" Then
                Index = 6
                Load FrmTela(X).Txt(FrmTela(X).Txt.Count)
                Set Ob = FrmTela(X).Txt(FrmTela(X).Txt.Count - 1)
            ElseIf Tipo = "Lst" Then
                Index = 7
                Load FrmTela(X).Lst(FrmTela(X).Lst.Count)
                Set Ob = FrmTela(X).Lst(FrmTela(X).Lst.Count - 1)
            End If
            FrmTela(X).O.Caption = Index
                
            Ob.Tag = Rs!Nome
            Ob.Appearance = Rs.Fields("3D")
            Ob.Text = Rs!Texto
            Ob.Stretch = Rs!Comprimir
            AbreFig Rs!Imagem, Ob, True
            Ob.BorderStyle = Rs!Borda
            Ob.Caption = Rs!Legenda
            Ob.Height = Rs!TamX
            Ob.Width = Rs!TamY
            Ob.BackColor = Rs.Fields("Cor Fundo")
            Ob.Top = Rs!PoxX
            Ob.Left = Rs!PoxY
            Ob.FontSize = Rs!Tamanho
            Ob.FontName = Rs!Estilo
            Ob.ForeColor = Rs.Fields("Cor da Letra")
            Ob.TabIndex = Rs!Ordem
            Ob.PasswordChar = Rs!Mascara
            Ob.Visible = True
            FrmCodigo.Eventos.Nodes.Add UCase(FrmTela(X).Tag), tvwChild, UCase(FrmTela(X).Tag + "." + Ob.Tag), Ob.Tag, 3

            With FrmCodigo.Eventos
            
                .Nodes.Add UCase(FrmTela(X).Tag + "." + Ob.Tag), tvwChild, UCase(FrmTela(X).Tag + "." + Ob.Tag + ".1"), "Ao Clicar 1 Vezes", 2
                If Index = 7 Or Index = 6 Or Index = 5 Or Index = 3 Or Index = 2 Or Index = 1 Then
                    .Nodes.Add UCase(FrmTela(X).Tag + "." + Ob.Tag), tvwChild, UCase(FrmTela(X).Tag + "." + Ob.Tag + ".2"), "Ao Clicar 2 Vezes", 2
                End If
                .Nodes.Add UCase(FrmTela(X).Tag + "." + Ob.Tag), tvwChild, UCase(FrmTela(X).Tag + "." + Ob.Tag + ".Ganhar"), "Ao Ganhar M focu", 2
                .Nodes.Add UCase(FrmTela(X).Tag + "." + Ob.Tag), tvwChild, UCase(FrmTela(X).Tag + "." + Ob.Tag + ".Perder"), "Ao Perder M focu", 2
                If Index = 6 Or 3 Then
                    .Nodes.Add UCase(FrmTela(X).Tag + "." + Ob.Tag), tvwChild, UCase(FrmTela(X).Tag + "." + Ob.Tag + ".Escrever"), "Ao Escrever", 2
                End If
                
            End With
            
            Xy = Xy + 1
            Rs.MoveNext
        Loop
        Set Rs = Banco.OpenRecordset("Select * From " & FrmTela(X).Tag & "", dbOpenDynaset)
        Do While Not Rs.EOF
            For Each Ob In FrmTela(X)
                If UCase(Rs!Nome) = UCase(Ob.Tag) Then
                    Ob.TabIndex = Rs!Ordem
                    Exit For
                End If
            Next
            Rs.MoveNext
        Loop
            
            
        
        Set Rs = Banco.OpenRecordset("Menu-" + FrmTela(X).Tag)
        If Rs.EOF = True Then GoTo ProximaTela:
        FrmTela(X).Visible = True
        FrmTela(X).Xmenu.ItemsFont = Rs!Legenda
        FrmTela(X).Xmenu.Style = Rs!Chave
        FrmTela(X).Xmenu.HighLightStyle = Rs!Root

        FrmTela(X).Show
        FrmTela(X).Refresh
        Rs.MoveNext
        If Rs.EOF = True Then GoTo ProximaTela
        
        Do While Not Rs.EOF
            
            Set M = New MenuItem
            M.Caption = Rs!Legenda
            M.Ident = Rs!Ident
            M.Name = Rs!Chave
            M.RootIndex = Rs!Root
            Nome = M.Caption
            InX = M.Ident
            M.Ident = InX
            M.Caption = Nome
            M.Name = Nome
            M.Accelerator = Nome
            M.Description = Nome
            If M.Ident = 0 Then
                FrmCodigo.Eventos.Nodes.Add UCase(FrmTela(TelaAtual).Tag), tvwChild, UCase(FrmTela(TelaAtual).Tag + "." + M.Caption + Trim(Str(M.Ident))), M.Caption, 5
            ElseIf M.Ident = 1 Then
                FrmCodigo.Eventos.Nodes.Add NomeMenu1, tvwChild, UCase(FrmTela(TelaAtual).Tag + "." + M.Caption + Trim(Str(M.Ident))), M.Caption, 5
            ElseIf M.Ident = 2 Then
                FrmCodigo.Eventos.Nodes.Add NomeMenu2, tvwChild, UCase(FrmTela(TelaAtual).Tag + "." + M.Caption + Trim(Str(M.Ident))), M.Caption, 5
            ElseIf M.Ident = 3 Then
                FrmCodigo.Eventos.Nodes.Add NomeMenu3, tvwChild, UCase(FrmTela(TelaAtual).Tag + "." + M.Caption + Trim(Str(M.Ident))), M.Caption, 5
            ElseIf M.Ident = 4 Then
                FrmCodigo.Eventos.Nodes.Add NomeMenu4, tvwChild, UCase(FrmTela(TelaAtual).Tag + "." + M.Caption + Trim(Str(M.Ident))), M.Caption, 5
            ElseIf M.Ident = 5 Then
                FrmCodigo.Eventos.Nodes.Add NomeMenu5, tvwChild, UCase(FrmTela(TelaAtual).Tag + "." + M.Caption + Trim(Str(M.Ident))), M.Caption, 5
            ElseIf M.Ident = 6 Then
                FrmCodigo.Eventos.Nodes.Add NomeMenu6, tvwChild, UCase(FrmTela(TelaAtual).Tag + "." + M.Caption + Trim(Str(M.Ident))), M.Caption, 5
            End If
            If M.Ident = 0 Then
                NomeMenu1 = UCase(FrmTela(TelaAtual).Tag + "." + M.Caption + Trim(Str(M.Ident)))
            ElseIf M.Ident = 1 Then
                NomeMenu2 = UCase(FrmTela(TelaAtual).Tag + "." + M.Name + Trim(Str(M.Ident)))
            ElseIf M.Ident = 2 Then
                NomeMenu3 = UCase(FrmTela(TelaAtual).Tag + "." + M.Name + Trim(Str(M.Ident)))
            ElseIf M.Ident = 3 Then
                NomeMenu4 = UCase(FrmTela(TelaAtual).Tag + "." + M.Name + Trim(Str(M.Ident)))
            ElseIf M.Ident = 4 Then
                NomeMenu5 = UCase(FrmTela(TelaAtual).Tag + "." + M.Name + Trim(Str(M.Ident)))
            ElseIf M.Ident = 5 Then
                NomeMenu6 = UCase(FrmTela(TelaAtual).Tag + "." + M.Name + Trim(Str(M.Ident)))
            End If
            FrmTela(X).Xmenu.MenuTree.Add M
            Rs.MoveNext
        Loop
        FrmTela(X).Xmenu.Refresh
ProximaTela:
        FrmTela(X).SetFocus
        TabelaIndex = TabelaIndex + 1
        X = X + 1
    Loop
    BuscaEventosArq ""
    Abilita True
End If
End Sub
