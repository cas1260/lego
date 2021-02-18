VERSION 5.00
Object = "{32A4927E-FB95-11D1-BF5B-00A024982E5B}#94.0#0"; "AXGRID.OCX"
Begin VB.Form Prop 
   BorderStyle     =   0  'None
   ClientHeight    =   5820
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4860
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   4860
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin axGridControl.axgrid Grind 
      Height          =   2355
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   4154
      Rows            =   0
      Cols            =   2
      Redraw          =   -1  'True
      ShowGrid        =   -1  'True
      GridSolid       =   0   'False
      GridLineColor   =   12632256
      BorderStyle     =   4
      BackColorFixed  =   16777215
      ColHeader       =   0   'False
      RowHeader       =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColorBkg    =   14737632
   End
End
Attribute VB_Name = "Prop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()
If Me.WindowState <> 1 Then
    Grind.Top = 0
    Grind.Left = 0
    Grind.Height = Me.ScaleHeight
    Grind.Width = Me.ScaleWidth
End If
End Sub
