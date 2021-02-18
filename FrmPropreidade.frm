VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form FrmPropreidade 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Propriedade"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   1905
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   1905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSDBGrid.DBGrid Pro 
      Bindings        =   "FrmPropreidade.frx":0000
      Height          =   1935
      Left            =   -450
      OleObjectBlob   =   "FrmPropreidade.frx":0014
      TabIndex        =   0
      Top             =   0
      Width           =   2385
   End
End
Attribute VB_Name = "FrmPropreidade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim V As Boolean

Private Sub Form_Load()
V = True
Form_Resize
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Cancel = 1
Me.Visible = False
End Sub

Private Sub Form_Resize()
If Me.WindowState <> 1 Then
    Me.Top = FrmPrincipal.Top + FrmPrincipal.Height
    Me.Left = FrmPrincipal.ScaleWidth - Me.ScaleWidth
    Me.Height = Screen.Height - Me.Top - 450
End If
If Me.WindowState = 1 Then
    V = FrmFerramentas.Visible
    FrmFerramentas.Visible = False
Else
    FrmFerramentas.WindowState = Me.WindowState
    FrmFerramentas.Visible = V
End If
End Sub

Private Sub Pro_KeyDown(KeyCode As Integer, Shift As Integer)
If Pro.Col = 0 Then Pro.Col = 1
End Sub
