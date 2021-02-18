VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmCodigo 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Codigo"
   ClientHeight    =   4905
   ClientLeft      =   4185
   ClientTop       =   1710
   ClientWidth     =   6345
   FillColor       =   &H00800000&
   Icon            =   "FrmCodigo.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4905
   ScaleWidth      =   6345
   Begin MSComctlLib.TreeView Eventos 
      Height          =   3195
      Left            =   240
      TabIndex        =   7
      Top             =   0
      Width           =   2865
      _ExtentX        =   5054
      _ExtentY        =   5636
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   529
      LabelEdit       =   1
      LineStyle       =   1
      PathSeparator   =   "."
      Sorted          =   -1  'True
      Style           =   7
      SingleSel       =   -1  'True
      ImageList       =   "Img3"
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ListBox Comandos 
      ForeColor       =   &H80000002&
      Height          =   840
      Left            =   420
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   1530
      Visible         =   0   'False
      Width           =   2625
   End
   Begin VB.ComboBox Lista 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   7140
      Index           =   0
      Left            =   0
      Sorted          =   -1  'True
      Style           =   1  'Simple Combo
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   2805
   End
   Begin VB.ComboBox Lista 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   7140
      Index           =   1
      Left            =   0
      Style           =   1  'Simple Combo
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   2805
   End
   Begin Lego.CodeHighlight TxtCod 
      Height          =   3372
      Left            =   3120
      TabIndex        =   3
      Top             =   504
      Width           =   2592
      _extentx        =   4577
      _extenty        =   5953
      operatorcolor   =   255
      delimitercolor  =   32768
      forecolor       =   -2147483643
      functioncolor   =   16711680
      font            =   "FrmCodigo.frx":030A
   End
   Begin MSComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   4590
      Width           =   6345
      _ExtentX        =   11192
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1411
            MinWidth        =   1411
            TextSave        =   "16:48"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "22/05/2004"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   630
      Top             =   3630
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCodigo.frx":0336
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCodigo.frx":078A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCodigo.frx":0D26
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCodigo.frx":1042
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCodigo.frx":1C96
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCodigo.frx":2572
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label index 
      Caption         =   "0"
      Height          =   465
      Left            =   1620
      TabIndex        =   4
      Top             =   2760
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label Label1 
      Height          =   435
      Left            =   2760
      TabIndex        =   1
      Top             =   3720
      Width           =   825
   End
End
Attribute VB_Name = "FrmCodigo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NomeLinha As String
Dim ListaChage As Boolean
Dim ConttxtCod As Long
Dim KeyCont As Long
Dim Impar_Par As Boolean
Private Sub Command1_Click()
Text1.Text = OrganizaSe(TxtCod.Text)
End Sub

Private Sub Eventos_NodeClick(ByVal Node As MSComctlLib.Node)
TxtCod.Text = Eventos.SelectedItem.Tag
Status.Panels(4).Text = Node.Text
End Sub
Private Sub Form_GotFocus()
If CRun.OpenExe = 0 Then FrmPrincipal.Visible = False
Set FrmCodigoRum = Me
End Sub

Private Sub Form_Initialize()
If CRun.OpenExe = 0 Then
    Me.Visible = False
    FrmPrincipal.Visible = False
    Exit Sub
End If

End Sub

Private Sub Form_Load()
'TxtCod.MultiLine = True
AddLista
Me.Width = 7305
Me.Top = 0
Me.Left = 0
Form_Resize
'FrmCodigo.Eventos.Nodes.Add , , "MOD", "Modulos", 4
'FrmCodigo.Eventos.Nodes.Add , , "PROD", "Procedimentos", 5
TxtCod.SelStart = 0
TxtCod.SelLength = Len(TxtCod.Text)
'TxtCod.TXTCOD.SelFontSize = 10
TxtCod.Language = 1
TxtCod.HighlightCode = 1
If CRun.OpenExe = 0 Then FrmPrincipal.Visible = False
If CRun.OpenExe = 0 Then Me.Visible = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Cancel = 1
Me.Visible = False
End Sub

Private Sub Form_Resize()
On Error Resume Next
If Me.WindowState <> 1 Then

    If CRun.OpenExe = 0 Then FrmPrincipal.Visible = False
    Eventos.Top = 0
    Eventos.Left = 0
    Eventos.Height = Me.ScaleHeight - Status.Height
    TxtCod.Left = Eventos.Width
    TxtCod.Top = 0
    TxtCod.Height = Me.Height - Status.Height - 350
    FrmCodigo.Lista(0).Top = Eventos.Top
    
    FrmCodigo.Lista(0).Left = Eventos.Left
    FrmCodigo.Lista(0).Width = Eventos.Width
    FrmCodigo.Lista(0).Height = Eventos.Height
    Lista(1).Top = Eventos.Top
    FrmCodigo.Lista(1).Left = Eventos.Left
    Lista(1).Width = Eventos.Width
    FrmCodigo.Lista(1).Height = Eventos.Height
    Status.Panels(4).Width = Me.Height
    If Me.ScaleWidth > 2820 Then
        TxtCod.Width = Me.ScaleWidth - Eventos.Width
    End If
    If CRun.OpenExe = 0 Then FrmPrincipal.Visible = False
    If CRun.OpenExe = 0 Then Me.Visible = False
End If
End Sub

Private Sub Label1_Change()
Status.Panels(3).Text = Label1.Caption
End Sub

Private Sub Label1_Click()
Status.Panels(3).Text = Label1.Caption
End Sub


Private Sub TxtCod_Change()
On Error Resume Next
FrmCodigo.Eventos.SelectedItem.Tag = TxtCod.Text
End Sub

Private Sub TxtCod_GotFocus()
Set FrmCodigoRum = Me
'TxtCod.SelColor = &H800000
'TxtCod.SelFontSize = 10
'TxtCod.SelFontName = "Courier New"
End Sub

Private Sub TxtCod_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 9 And Shift = 0 Then
    KeyCode = 0
    TxtCod.SelText = "    "
End If
If KeyCode = 32 And Shift = 2 Then
    TipoMenu = False
    FrmCodigo.Lista(0).Visible = True
    FrmCodigo.Lista(0).SetFocus
    KeyCode = 0
End If
If KeyCode = 32 And Shift = 3 Then
    AddObj
    TipoMenu = True
    FrmCodigo.Lista(1).Visible = True
    FrmCodigo.Lista(1).SetFocus
    KeyCode = 0
End If

End Sub


Public Sub AddLista()
On Error Resume Next
Dim X As Long
FrmCodigo.Lista(0).AddItem "Legenda"
FrmCodigo.Lista(0).AddItem "Cor de Fundo"
FrmCodigo.Lista(0).AddItem "Cor da Letra"
FrmCodigo.Lista(0).AddItem "TamX"
FrmCodigo.Lista(0).AddItem "TamY"
FrmCodigo.Lista(0).AddItem "PoxX"
FrmCodigo.Lista(0).AddItem "PoxY"
FrmCodigo.Lista(0).AddItem "Fonte"
FrmCodigo.Lista(0).AddItem "Tamanho"
FrmCodigo.Lista(0).AddItem "Texto"
FrmCodigo.Lista(0).AddItem "Imagem"
FrmCodigo.Lista(0).AddItem "Borda"
FrmCodigo.Lista(0).AddItem "Comprimir"
FrmCodigo.Lista(0).AddItem "Ordem"
FrmCodigo.Lista(0).AddItem "3d"
FrmCodigo.Lista(0).AddItem "Marca"
FrmCodigo.Lista(0).AddItem "Mascara"
FrmCodigo.Lista(0).AddItem "Total"
FrmCodigo.Lista(0).AddItem "Atual"
FrmCodigo.Lista(0).AddItem "Busca"

FrmCodigo.Lista(0).AddItem "Va Para"
FrmCodigo.Lista(0).AddItem "Se"
FrmCodigo.Lista(0).AddItem "SeNao"
FrmCodigo.Lista(0).AddItem "Msg"
FrmCodigo.Lista(0).AddItem "Loop "
FrmCodigo.Lista(0).AddItem "Fim do Programa"
FrmCodigo.Lista(0).AddItem "FimSe"
FrmCodigo.Lista(0).AddItem "Chama"
FrmCodigo.Lista(0).AddItem "Fecha Tela"
FrmCodigo.Lista(0).AddItem "Focus"
FrmCodigo.Lista(0).AddItem "Adicione"
FrmCodigo.Lista(0).AddItem "Selecione"
FrmCodigo.Lista(0).AddItem "Limpa"
FrmCodigo.Lista(0).AddItem "SimNao"
FrmCodigo.Lista(0).AddItem "Exe "
FrmCodigo.Lista(0).AddItem "Direita"
FrmCodigo.Lista(0).AddItem "Esquerda"
FrmCodigo.Lista(0).AddItem "Procedimento"
FrmCodigo.Lista(0).AddItem "Db "
FrmCodigo.Lista(0).AddItem "Tela Cheia()"
FrmCodigo.Lista(0).AddItem "Caixa"
FrmCodigo.Lista(0).AddItem "Grava("
FrmCodigo.Lista(0).AddItem "Deletar("
FrmCodigo.Lista(0).AddItem "LimpaArquivo"
FrmCodigo.Lista(0).AddItem "Faça Enquanto "
FrmCodigo.Lista(0).AddItem "Visivel"
If CRun.OpenExe = 0 Then FrmPrincipal.Visible = False
End Sub

Private Sub TxtCod_KeyPress(KeyAscii As Integer)
If KeyAscii = Asc(".") Then
    If BuscaPonto() = True Then
        TipoMenu = True
        FrmCodigo.Lista(0).Visible = True
        FrmCodigo.Lista(0).ZOrder vbBringToFront
        FrmCodigo.Lista(0).SetFocus
        KeyAscii = 0
    End If
ElseIf KeyAscii = Asc(";") Then
    AddObj
    TipoMenu = True
    If FrmCodigo.Lista(1).ListCount <> 0 Then
        FrmCodigo.Lista(1).Visible = True
        FrmCodigo.Lista(1).ZOrder vbBringToFront
        FrmCodigo.Lista(1).SetFocus
        KeyAscii = 0
    End If
ElseIf KeyAscii = Asc("(") Then
    TxtCod.SelText = "()"
    TxtCod.SelStart = TxtCod.SelStart - 1
    KeyAscii = 0
ElseIf KeyAscii = Asc("{") Then
    TxtCod.SelText = "{}"
    TxtCod.SelStart = TxtCod.SelStart - 1
    KeyAscii = 0
ElseIf KeyAscii = Asc("[") Then
    TxtCod.SelText = "[]"
    TxtCod.SelStart = TxtCod.SelStart - 1
    KeyAscii = 0
ElseIf KeyAscii = 34 Then
    TxtCod.SelText = Chr(34) + Chr(34)
    TxtCod.SelStart = TxtCod.SelStart - 1
    KeyAscii = 0


End If

End Sub

Private Sub TxtCod_SelChange()
GetEditStatus TxtCod.TxtCod, Label1
Status.Panels(4).Text = TxtCod.Line(TxtCod.LineIndex)
'MsgBox TxtCod.SelCharOffset
End Sub



Public Sub AddObj()
On Error Resume Next
Dim X As Long, Nome As String, Nome2 As String
Dim Obj As Object, Y As Long

FrmCodigo.Lista(1).Clear
For X = 0 To ContTela - 1
    Nome = FrmTela(X).Tag
    Y = InStr(1, FrmCodigo.Eventos.SelectedItem.Key, ".")
    
    If Y = 0 Then
        Nome2 = FrmCodigo.Eventos.SelectedItem.Key
    Else
        Nome2 = Left(FrmCodigo.Eventos.SelectedItem.Key, Y - 1)
    End If
    If UCase(Nome) <> UCase(Nome2) Then
        GoTo proximo
    End If
    
    For Each Obj In FrmTela(X)
        If Obj.Name = "xMenu" Or Obj.Name = "Im" Or Obj.Name = "Focus" Or Obj.Name = "S" Or Obj.Name = "Nome" Or Obj.Name = "Cont" Then
            
        ElseIf Obj.index = 0 Then
            
        Else
            Nome = Obj.Tag
            FrmCodigo.Lista(1).AddItem Nome
        End If
    Next
proximo:
Next X

End Sub
Private Function BuscaPonto() As Boolean
On Error Resume Next
Dim Com As String
Dim ContX As Long, X As Long
Dim Pos As Long

Com = TxtCod.Line(TxtCod.LineIndex)
ContX = 0
For X = 1 To Len(Com)
    If Mid(Com, X, 1) = Chr(34) Then
        ContX = ContX + 1
    End If
Next X
Pos = Right(Str(ContX), 1)
Select Case Pos
    Case 0, 1, 3, 5, 7, 9
        BuscaPonto = True
    Case 2, 4, 6, 8
        BuscaPonto = False
End Select
End Function

Public Sub Fechar(index As Integer)
If TipoMenu = True Then
    If index = 0 Then
        TxtCod.SelText = "." + Lista(index).Text
    Else
        TxtCod.SelText = ";" + Lista(index).Text
    End If
Else
    TxtCod.SelText = Lista(index).Text
End If
TxtCod.SetFocus

End Sub

Private Sub Lista_Change(index As Integer)
If ListaChage = True Then
    SelectComboBox index
End If
ListaChage = True
End Sub

Private Sub Lista_DblClick(index As Integer)
Fechar index
End Sub

Private Sub Lista_GotFocus(index As Integer)
Eventos.Visible = False
If Lista(index).ListCount <> -1 And Lista(index).ListCount <> 0 Then
    Lista(index).ListIndex = 0
End If
End Sub

Private Sub Lista_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Fechar index
ElseIf KeyCode = 27 Then
    Lista_LostFocus index
    FrmCodigoRum.TxtCod.SetFocus
    If index = 0 Then
        FrmCodigoRum.TxtCod.SelText = "."
    End If
ElseIf KeyCode = vbKeyBack Then
    ListaChage = False
    Lista(index).SelText = ""
    If Len(Lista(index).Text) > 0 Then
        Lista(index).Text = Left(Lista(index).Text, Len(Lista(index).Text) - 1)
    End If
    ListaChage = True
    KeyCode = 0

End If
End Sub

Private Sub Lista_KeyPress(index As Integer, KeyAscii As Integer)
If KeyAscii = Asc(".") Then
    If index = 1 Then
        Fechar index
        KeyAscii = 0
        Lista(0).Visible = True
        Lista(0).SetFocus
    End If
End If
End Sub

Private Sub Lista_LostFocus(index As Integer)
Eventos.Visible = True
Lista(index).Visible = False
End Sub


Private Sub SelectComboBox(index As Integer)
On Error Resume Next
Dim Nome As String, X As Long

Nome = Trim(Lista(index).Text)
If Trim(Nome) = "" Then
    Exit Sub
End If
For X = 0 To Lista(index).ListCount
    If UCase(Nome) = Trim(UCase(Left(Lista(index).List(X), Len(Nome)))) Then
        'Lista.ListIndex = X
        GoTo Fim
    End If
Next X

Exit Sub

Fim:

Lista(index).Text = Lista(index).List(X)
Lista(index).SelStart = Len(Nome)
Lista(index).SelLength = Len(Lista(index).List(X)) - Len(Nome)

End Sub
