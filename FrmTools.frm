VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmTools 
   Caption         =   "Ferramentas"
   ClientHeight    =   6645
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3195
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   6645
   ScaleWidth      =   3195
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox ListaProg 
      Height          =   1035
      Left            =   2970
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   6240
      Visible         =   0   'False
      Width           =   1260
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3504
      Top             =   2508
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483624
      ImageWidth      =   28
      ImageHeight     =   28
      MaskColor       =   16776960
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTools.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTools.frx":0984
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTools.frx":1308
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTools.frx":1C8C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTools.frx":2610
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTools.frx":2F94
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTools.frx":3918
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTools.frx":429C
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTools.frx":4C20
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTools.frx":55A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTools.frx":5E80
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTools.frx":675C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar T 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   3195
      _ExtentX        =   5636
      _ExtentY        =   1005
      ButtonWidth     =   926
      ButtonHeight    =   900
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   12
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
            Key             =   "I"
            Object.ToolTipText     =   "Banco de Dados"
            ImageIndex      =   10
            Style           =   1
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "L"
            Object.ToolTipText     =   "Tabelas (RecordSets)"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "M"
            Object.ToolTipText     =   "Timer"
            ImageIndex      =   12
         EndProperty
      EndProperty
      Enabled         =   0   'False
   End
   Begin MSComctlLib.TreeView Eventos 
      Height          =   3255
      Left            =   0
      TabIndex        =   2
      Top             =   1080
      Width           =   3195
      _ExtentX        =   5636
      _ExtentY        =   5741
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   529
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      SingleSel       =   -1  'True
      ImageList       =   "ImageList2"
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
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   0
      Top             =   0
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
            Picture         =   "FrmTools.frx":6EB0
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTools.frx":7304
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTools.frx":78A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTools.frx":7BBC
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTools.frx":8810
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTools.frx":90EC
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmTools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.Top = FrmPrincipal.Height
Me.Left = 0
Me.Height = Screen.Height - FrmPrincipal.Height - 400
FrmTools.Eventos.Nodes.Add , , "MOD", "Modulos", 4
FrmTools.Eventos.Nodes.Add , , "PROD", "Procedimentos", 5
End Sub

Private Sub Form_Resize()
If Me.WindowState <> 1 Then
    Eventos.Height = Me.ScaleHeight - T.Height
End If
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
Private Sub Eventos_NodeClick(ByVal Node As MSComctlLib.Node)
Dim Frm1 As New FrmCodigo
Frm1.P.Caption = Eventos.SelectedItem.Index
Frm1.TxtCod.Text = Eventos.SelectedItem.Tag
Frm1.Show
Frm1.Status.Panels(4).Text = Node.Text
End Sub
