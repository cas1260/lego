VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Analizador de expresiones matemáticas"
   ClientHeight    =   945
   ClientLeft      =   1185
   ClientTop       =   1800
   ClientWidth     =   6105
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   945
   ScaleWidth      =   6105
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   4650
      TabIndex        =   2
      Top             =   540
      Width           =   705
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Calcular"
      Height          =   375
      Left            =   4440
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Text            =   "3 + tan( pi / 4 )"
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim hParse As Long ' handle




Private Sub Command1_Click()
 Dim d As Double
 Dim n As Integer
 Dim e As String * 64
 If GetValue(0, Text1, d) Then
   Text1 = d
 Else
   n = GetError(hParse, e)
   Text1 = "Error " & n & " :   " & e
 End If
End Sub


Private Sub Command2_Click()

SetValue 0, "Tipo1", 250
SetValueTxt 0, "Tipo", "150"
MsgBox GetValue(0, "Tipo1", "0")

End Sub

Private Sub Form_Load()
  hParse = ParseInit
End Sub


Private Sub Form_Unload(Cancel As Integer)
  ParseEnd hParse
End Sub


