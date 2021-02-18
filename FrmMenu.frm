VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmMenu 
   Caption         =   "Editor de Menu - Lego 1.1"
   ClientHeight    =   4260
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5820
   ClipControls    =   0   'False
   Icon            =   "FrmMenu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   5820
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab SSTab1 
      Height          =   4245
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   5805
      _ExtentX        =   10239
      _ExtentY        =   7488
      _Version        =   393216
      TabOrientation  =   2
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Geral"
      TabPicture(0)   =   "FrmMenu.frx":0442
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(2)=   "Fontes"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Menu"
      TabPicture(1)   =   "FrmMenu.frx":045E
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Incluir"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Command1"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "TxtLeg"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Command4"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Setas(0)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Setas(1)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Setas(2)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Setas(3)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Command2"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "ListMenu"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).ControlCount=   11
      Begin VB.ListBox ListMenu 
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2790
         Left            =   480
         TabIndex        =   7
         Top             =   1020
         Width           =   5145
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Cancelar"
         Height          =   345
         Left            =   4530
         TabIndex        =   27
         Top             =   3840
         Width           =   1125
      End
      Begin VB.CommandButton Setas 
         Height          =   405
         Index           =   3
         Left            =   2370
         Picture         =   "FrmMenu.frx":047A
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   570
         Width           =   645
      End
      Begin VB.CommandButton Setas 
         Height          =   405
         Index           =   2
         Left            =   1740
         Picture         =   "FrmMenu.frx":08BC
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   570
         Width           =   645
      End
      Begin VB.CommandButton Setas 
         Height          =   405
         Index           =   1
         Left            =   1110
         Picture         =   "FrmMenu.frx":0CFE
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   570
         Width           =   645
      End
      Begin VB.CommandButton Setas 
         Height          =   405
         Index           =   0
         Left            =   480
         Picture         =   "FrmMenu.frx":1140
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   570
         Width           =   645
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&Delete"
         Height          =   345
         Left            =   4500
         TabIndex        =   6
         Top             =   570
         Width           =   1125
      End
      Begin VB.Frame Fontes 
         Caption         =   "Fontes"
         ForeColor       =   &H00FF0000&
         Height          =   1875
         Left            =   -74550
         TabIndex        =   24
         Top             =   2070
         Width           =   5145
         Begin VB.ListBox ListFonts 
            ForeColor       =   &H00FF0000&
            Height          =   840
            ItemData        =   "FrmMenu.frx":1582
            Left            =   120
            List            =   "FrmMenu.frx":1589
            Sorted          =   -1  'True
            TabIndex        =   18
            Top             =   510
            Width           =   4935
         End
         Begin VB.Label L 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Teste de Fontes"
            ForeColor       =   &H00800000&
            Height          =   435
            Left            =   120
            TabIndex        =   26
            Top             =   1380
            Width           =   4950
         End
         Begin VB.Label Label3 
            Caption         =   "Nome da Fonte"
            ForeColor       =   &H00FF0000&
            Height          =   225
            Left            =   120
            TabIndex        =   25
            Top             =   270
            Width           =   1215
         End
      End
      Begin VB.TextBox TxtLeg 
         Height          =   285
         Left            =   1290
         TabIndex        =   0
         Top             =   150
         Width           =   4365
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Ok"
         Height          =   345
         Left            =   3420
         TabIndex        =   22
         Top             =   3840
         Width           =   1125
      End
      Begin VB.CommandButton Incluir 
         Caption         =   "&Incluir"
         Default         =   -1  'True
         Height          =   345
         Left            =   3390
         TabIndex        =   5
         Top             =   570
         Width           =   1125
      End
      Begin VB.Frame Frame2 
         Caption         =   "Seleção"
         ForeColor       =   &H00FF0000&
         Height          =   1155
         Left            =   -74550
         TabIndex        =   21
         Top             =   900
         Width           =   5115
         Begin VB.OptionButton Cel 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Caption         =   "3 D Suave"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   5
            Left            =   3300
            TabIndex        =   17
            Top             =   690
            Width           =   1245
         End
         Begin VB.OptionButton Cel 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Caption         =   "3 D Metalica"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   4
            Left            =   1830
            TabIndex        =   16
            Top             =   690
            Width           =   1245
         End
         Begin VB.OptionButton Cel 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Caption         =   "Gradiente Azul"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   3
            Left            =   90
            TabIndex        =   15
            Top             =   660
            Width           =   1425
         End
         Begin VB.OptionButton Cel 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Caption         =   "Gradiente Cinza"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   2
            Left            =   3300
            TabIndex        =   14
            Top             =   360
            Width           =   1515
         End
         Begin VB.OptionButton Cel 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Caption         =   "Solido"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   1
            Left            =   1830
            TabIndex        =   13
            Top             =   360
            Width           =   765
         End
         Begin VB.OptionButton Cel 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Caption         =   "Simples"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   0
            Left            =   90
            TabIndex        =   12
            Top             =   360
            Value           =   -1  'True
            Width           =   855
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Aparencia"
         ForeColor       =   &H00FF0000&
         Height          =   795
         Left            =   -74580
         TabIndex        =   20
         Top             =   90
         Width           =   5145
         Begin VB.OptionButton Aparencia 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Caption         =   "Office 97"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   8
            Top             =   420
            Value           =   -1  'True
            Width           =   945
         End
         Begin VB.OptionButton Aparencia 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Caption         =   "Office Xp"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   1
            Left            =   1200
            TabIndex        =   9
            Top             =   420
            Width           =   1155
         End
         Begin VB.OptionButton Aparencia 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Caption         =   "Grafico"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   2
            Left            =   2400
            TabIndex        =   10
            Top             =   420
            Width           =   1005
         End
         Begin VB.OptionButton Aparencia 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Caption         =   "3 D"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   3
            Left            =   3600
            TabIndex        =   11
            Top             =   390
            Width           =   945
         End
      End
      Begin VB.Label Label1 
         Caption         =   "Legenda :"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   510
         TabIndex        =   23
         Top             =   210
         Width           =   735
      End
   End
End
Attribute VB_Name = "FrmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Tipo As String

Private Sub Command1_Click()
On Error Resume Next
Dim X As Long, Nome As String, Y As Long
Dim O As New MenuItem, InX As Long
Dim NovoI As String
Dim NomeMenu1 As String, NomeMenu2 As String
Dim NomeMenu3 As String, NomeMenu4 As String
Dim NomeMenu5 As String, NomeMenu6 As String
Dim NomeTag() As String, ContTag() As String
Dim ContMenu As Long, Totalmenu As String

X = FrmTela(TelaAtual).Xmenu.MenuTree.Count
ReDim NomeTag(X) As String, ContTag(X) As String

ContMenu = 0

Do While X <> 0
    For Y = 1 To FrmCodigo.Eventos.Nodes.Count
        NovoI = UCase(FrmTela(TelaAtual).Tag + "." + UCase(FrmTela(TelaAtual).Xmenu.MenuTree(X).Caption)) + Trim(Str(FrmTela(TelaAtual).Xmenu.MenuTree(X).Ident))
        If UCase(FrmCodigo.Eventos.Nodes(Y).Key) = NovoI Then
            NomeTag(ContMenu) = NovoI
            ContTag(ContMenu) = FrmCodigo.Eventos.Nodes(Y).Tag
            ContMenu = ContMenu + 1
            FrmCodigo.Eventos.Nodes.Remove Y
            Exit For
        End If
    Next Y
    FrmTela(TelaAtual).Xmenu.MenuTree.Remove X
    X = X - 1
Loop

Totalmenu = ContMenu
InX = 0
For X = 0 To ListMenu.ListCount - 1
    Nome = ListMenu.List(X)
    If Trim(Nome) = "" Or Nome = ">" Or Nome = ">>" Or Nome = ">>>>" Or Nome = ">>>>>" Or Nome = ">>>>>" Then
        X = X + 1
        Exit For
    End If
    Y = 6
    Do While Y <> 0
        If Left(Nome, Y) = String(Y, ">") Then
            InX = Y
            Nome = Right(Nome, Len(Nome) - Y)
            Exit Do
        End If
        Y = Y - 1
    Loop
    InX = Y
    O.Ident = InX
    O.Caption = Nome
    O.Name = Nome
    O.Accelerator = Nome
    O.Description = Nome
    FrmTela(TelaAtual).Xmenu.MenuTree.Add O
    If O.Ident = 0 Then
        FrmCodigo.Eventos.Nodes.Add UCase(FrmTela(TelaAtual).Tag), tvwChild, UCase(FrmTela(TelaAtual).Tag + "." + O.Caption + Trim(Str(O.Ident))), O.Caption, 5
    ElseIf O.Ident = 1 Then
        FrmCodigo.Eventos.Nodes.Add NomeMenu1, tvwChild, UCase(FrmTela(TelaAtual).Tag + "." + O.Caption + Trim(Str(O.Ident))), O.Caption, 5
    ElseIf O.Ident = 2 Then
        FrmCodigo.Eventos.Nodes.Add NomeMenu2, tvwChild, UCase(FrmTela(TelaAtual).Tag + "." + O.Caption + Trim(Str(O.Ident))), O.Caption, 5
    ElseIf O.Ident = 3 Then
        FrmCodigo.Eventos.Nodes.Add NomeMenu3, tvwChild, UCase(FrmTela(TelaAtual).Tag + "." + O.Caption + Trim(Str(O.Ident))), O.Caption, 5
    ElseIf O.Ident = 4 Then
        FrmCodigo.Eventos.Nodes.Add NomeMenu4, tvwChild, UCase(FrmTela(TelaAtual).Tag + "." + O.Caption + Trim(Str(O.Ident))), O.Caption, 5
    ElseIf O.Ident = 5 Then
        FrmCodigo.Eventos.Nodes.Add NomeMenu5, tvwChild, UCase(FrmTela(TelaAtual).Tag + "." + O.Caption + Trim(Str(O.Ident))), O.Caption, 5
    ElseIf O.Ident = 6 Then
        FrmCodigo.Eventos.Nodes.Add NomeMenu6, tvwChild, UCase(FrmTela(TelaAtual).Tag + "." + O.Caption + Trim(Str(O.Ident))), O.Caption, 5
    End If
        
    If O.Ident = 0 Then
        NomeMenu1 = UCase(FrmTela(TelaAtual).Tag + "." + O.Caption + Trim(Str(O.Ident)))
        NovoI = NomeMenu1
    ElseIf O.Ident = 1 Then
        NomeMenu2 = UCase(FrmTela(TelaAtual).Tag + "." + O.Name + Trim(Str(O.Ident)))
        NovoI = NomeMenu2
    ElseIf O.Ident = 2 Then
        NomeMenu3 = UCase(FrmTela(TelaAtual).Tag + "." + O.Name + Trim(Str(O.Ident)))
        NovoI = NomeMenu3
    ElseIf O.Ident = 3 Then
        NomeMenu4 = UCase(FrmTela(TelaAtual).Tag + "." + O.Name + Trim(Str(O.Ident)))
        NovoI = NomeMenu4
    ElseIf O.Ident = 4 Then
        NomeMenu5 = UCase(FrmTela(TelaAtual).Tag + "." + O.Name + Trim(Str(O.Ident)))
        NovoI = NomeMenu5
    ElseIf O.Ident = 5 Then
        NomeMenu6 = UCase(FrmTela(TelaAtual).Tag + "." + O.Name + Trim(Str(O.Ident)))
        NovoI = NomeMenu6
    End If
    
    Set O = New MenuItem
   
Next X
ContMenu = 0
For ContMenu = 0 To Totalmenu
    For XU = 1 To FrmCodigo.Eventos.Nodes.Count
        If UCase(FrmCodigo.Eventos.Nodes(XU).Key) = UCase(NomeTag(ContMenu)) Then
            FrmCodigo.Eventos.Nodes(XU).Tag = ContTag(ContMenu)
            Exit For
        End If
    Next XU
Next ContMenu

Set O = Nothing
Dim Apa As Long
If Aparencia(0).Value = True Then
    Apa = 0
ElseIf Aparencia(1).Value = True Then
     Apa = 1
ElseIf Aparencia(2).Value = True Then
    Apa = 2
ElseIf Aparencia(3).Value = True Then
    Apa = 3
Else
    Apa = 0
End If

FrmTela(TelaAtual).Xmenu.Style = Apa

If Cel(0).Value = True Then
    Apa = 0
ElseIf Cel(1).Value = True Then
    Apa = 1
ElseIf Cel(2).Value = True Then
    Apa = 2
ElseIf Cel(3).Value = True Then
    Apa = 3
ElseIf Cel(4).Value = True Then
    Apa = 4
ElseIf Cel(5).Value = True Then
    Apa = 5
End If
FrmTela(TelaAtual).Xmenu.ItemsFont = ListFonts.List(ListFonts.ListIndex)
FrmTela(TelaAtual).Xmenu.HighLightStyle = Apa
FrmTela(TelaAtual).Xmenu.Refresh

Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command4_Click()
If ListMenu.ListIndex <> -1 Then
    ListMenu.RemoveItem ListMenu.ListIndex
End If
End Sub

Private Sub Incluir_Click()
On Error Resume Next
If Trim(TxtLeg.Text) = "" Then
    SSTab1.Tab = 1
    TxtLeg.SetFocus
    Exit Sub
End If
FrmTela(TelaAtual).Xmenu.Visible = True
ListMenu.List(ListMenu.ListCount - 1) = Tipo + TxtLeg.Text
ListMenu.AddItem ""
ListMenu.Selected(ListMenu.ListCount - 1) = True
TxtLeg.Text = ""
TxtLeg.SetFocus
End Sub

Private Sub Form_Load()
Dim X As Long
ListFonts.Clear
For X = 0 To Screen.FontCount - 1
    ListFonts.AddItem Screen.Fonts(X)
Next X
Abrir
Tipo = ""
End Sub
Private Sub ListFonts_Click()
If ListFonts.ListIndex <> -1 Then
    L.FontName = ListFonts.List(ListFonts.ListIndex)
End If
End Sub

Private Sub ListTamanho_Click()
ListFonts_Click
End Sub

Private Sub Abrir()
Dim X As Long, Apa As Long
With FrmTela(TelaAtual).Xmenu
    For X = 1 To .MenuTree.Count
        ListMenu.AddItem String(.MenuTree(X).Ident, ">") + .MenuTree(X).Caption
    Next X
    ListMenu.AddItem ""
    ListMenu.Selected(ListMenu.ListCount - 1) = True
    Aparencia(.Style).Value = True
    Cel(.HighLightStyle).Value = True
    ListFonts.Text = FrmTela(TelaAtual).Xmenu.ItemsFont
End With
End Sub

Private Sub ListMenu_Click()
If ListMenu.ListIndex <> -1 Then
    TxtLeg.Text = ListMenu.List(ListMenu.ListIndex)
End If
End Sub

Private Sub Setas_Click(index As Integer)
Dim Nome1 As String, X As Long
Dim NomeAnt As String, NomeNovo As String, Cap As String, CapNovo As String
If ListMenu.ListIndex = -1 Then
    Exit Sub
End If

Nome1 = ListMenu.List(ListMenu.ListIndex)

Select Case index
    Case 0
        If Left(Nome1, 1) = ">" Then
            ListMenu.List(ListMenu.ListIndex) = Right(Nome1, Len(Nome1) - 1)
        End If
        If Len(Tipo) <> 0 Then
            Tipo = Left(Tipo, Len(Tipo) - 1)
        End If
    Case 1
        ListMenu.List(ListMenu.ListIndex) = ">" + Nome1
        Tipo = Tipo + ">"
    Case 2
        If ListMenu.ListIndex > 0 Then
            X = ListMenu.ListIndex - 1
            NomeAnt = ListMenu.List(X)
            NomeNovo = ListMenu.List(ListMenu.ListIndex)
            
            ListMenu.List(X) = NomeNovo
            ListMenu.List(X + 1) = NomeAnt
            ListMenu.Selected(X) = True
        End If
    Case 3
        If ListMenu.ListIndex < ListMenu.ListCount Then
            X = ListMenu.ListIndex + 1
            If ListMenu.ListIndex = ListMenu.ListCount - 1 Then
                Exit Sub
            End If
            NomeAnt = ListMenu.List(X)
            NomeNovo = ListMenu.List(ListMenu.ListIndex)
           
            ListMenu.List(X) = NomeNovo
            ListMenu.List(X - 1) = NomeAnt
            ListMenu.Selected(X) = True
        End If
End Select
TxtLeg.SetFocus
End Sub

Private Sub TxtLeg_Change()
If ListMenu.ListIndex <> -1 Then
    ListMenu.List(ListMenu.ListIndex) = Tipo + TxtLeg.Text
End If
End Sub
Private Function Elimina(Key As String)
On Error Resume Next
Dim X As Long, Inicio As Long
Dim NovaVar As String, Posicao() As String
Dim Xy As Long, Part As String
ReDim Posicao(11) As String
Dim Antiga As String
Dim AntigaVar As String
Dim Var1 As String

Xy = 0
Var = Key
AntigaVar = Var
Var1 = Var
NovaVar = Var
Inicio1 = 1
Antiga = ""

Dentro = True

For Xy = 1 To Len(Key)
    Var1 = Asc(Mid(Var, Xy, 1))
    X = Asc(UCase(Mid(Var, Xy, 1)))
    If X >= vbKeyA And X <= vbKeyZ Then
        Antiga = Antiga + Chr(Var1)
    ElseIf X >= vbKey0 And X <= vbKey9 Then
        Antiga = Antiga + Chr(Var1)
    Else
        If Antiga <> "" Then
            If Dentro = True Then
                AntigaVar = Antiga
                NovaVar = Replace(NovaVar, Antiga, AntigaVar)
                Antiga = ""
            Else
                Antiga = Antiga + Chr(Var1)
            End If
        End If
    End If
Next

Elimina = NovaVar
End Function
