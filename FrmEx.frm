VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmEx 
   Caption         =   "Projeto"
   ClientHeight    =   4104
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   3420
   Icon            =   "FrmEx.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4104
   ScaleWidth      =   3420
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1650
      Top             =   3510
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEx.frx":0E42
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEx.frx":1C96
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEx.frx":20EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEx.frx":2F3E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEx.frx":3392
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView Prog 
      Height          =   3048
      Left            =   216
      TabIndex        =   0
      Top             =   912
      Width           =   2832
      _ExtentX        =   4995
      _ExtentY        =   5376
      _Version        =   393217
      Indentation     =   159
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      FullRowSelect   =   -1  'True
      SingleSel       =   -1  'True
      ImageList       =   "ImageList1"
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "FrmEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
If CRun.OpenExe = 0 Then FrmPrincipal.Visible = False
End Sub

Private Sub Form_Load()

Me.Height = (FrmPrincipal.ScaleHeight / 2) - 100
Me.Width = 3360
FrmEx.Left = FrmPrincipal.ScaleWidth - FrmEx.Width
FrmEx.Top = 0
If CRun.OpenExe = 0 Then FrmPrincipal.Visible = False
End Sub
Private Sub Form_Unload(Cancel As Integer)
Cancel = 1
Me.Visible = False
End Sub
Private Sub Form_Resize()
If Me.WindowState <> 1 Then
    Prog.Top = 0
    Prog.Left = 0
    Prog.Height = Me.ScaleHeight
    Prog.Width = Me.ScaleWidth
    If CRun.OpenExe = 0 Then FrmPrincipal.Visible = False
End If
End Sub

Private Sub Prog_DblClick()
Dim a As String, X As Long
If Prog.Nodes.Count <> 0 Then
    a = Prog.SelectedItem
    If Trim(a) <> "" Then
        For X = 0 To ContTela - 1
            If UCase(a) = UCase(FrmTela(X).Tag) Then
                Set NovoObj = FrmTela(X)
                TelaAtual = X
                FrmTela(X).Visible = True
                FrmTela(X).SetFocus
                Exit Sub
            End If
        Next X
    End If
End If
End Sub

Private Sub Prog_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub
