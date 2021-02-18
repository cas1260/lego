VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl UserControl1 
   BackColor       =   &H8000000B&
   ClientHeight    =   2652
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3888
   ScaleHeight     =   2652
   ScaleWidth      =   3888
   Tag             =   "a"
   ToolboxBitmap   =   "UserControl11.ctx":0000
   Begin MSComctlLib.ImageList ImageList1 
      Index           =   1
      Left            =   600
      Top             =   2040
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Index           =   0
      Left            =   1320
      Top             =   2040
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   1920
      ScaleHeight     =   612
      ScaleWidth      =   1236
      TabIndex        =   0
      Top             =   360
      Width           =   1235
      Begin VB.Shape Shape1 
         Height          =   615
         Left            =   0
         Top             =   0
         Width           =   1215
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   2
         Index           =   5
         Visible         =   0   'False
         X1              =   1180
         X2              =   1180
         Y1              =   0
         Y2              =   600
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   2
         Index           =   4
         Visible         =   0   'False
         X1              =   0
         X2              =   1200
         Y1              =   570
         Y2              =   570
      End
      Begin VB.Image Image1 
         Height          =   375
         Index           =   0
         Left            =   60
         Top             =   60
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Index           =   3
         X1              =   0
         X2              =   1200
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   0
         X1              =   10
         X2              =   10
         Y1              =   0
         Y2              =   600
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   1
         X1              =   0
         X2              =   1200
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Index           =   2
         X1              =   1220
         X2              =   1220
         Y1              =   0
         Y2              =   600
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Command"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   180
         Width           =   735
      End
   End
   Begin VB.Image Image1 
      Height          =   375
      Index           =   1
      Left            =   2880
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   375
      Index           =   2
      Left            =   2880
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   495
   End
End
Attribute VB_Name = "UserControl1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim Control As Object

Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Event click()
Event dobleclick()


Public Enum aAlingment
LeftJustify = 0
RightJustify = 1
Center = 2
End Enum

Public Enum aborderstyle
None = 0
Fixed = 1
Raised = 2
End Enum

Public Enum aappearance
Flat = 0
cc3D = 1

End Enum

Public Enum avalue
Defaulf = 0
vbChecked = 1
End Enum

Public Enum aStyle
Standard = 0
Graphical = 1
End Enum

Dim u
Dim i, j As Integer
Dim icaption As String
Dim ienabled As Boolean
Dim iborderstyle As aborderstyle
Dim iappearance As aappearance
Dim iStyle As aStyle
Dim ibackcolor As OLE_COLOR
Dim iFont As IFontDisp
Dim ipicture As StdPicture
Dim ivalue As avalue
Dim ialingment As aAlingment
Dim Idisabledpicture As StdPicture
Dim iforecolor As OLE_COLOR
Public Property Get Enabled() As Boolean
Enabled = ienabled
End Property
Public Property Let Enabled(ByVal new_enabled As Boolean)
ienabled = new_enabled
UserControl.Enabled = new_enabled
PropertyChanged "Enabled"
End Property

Public Property Get Font() As IFontDisp
Set Font = iFont
End Property
Public Property Set Font(ByVal new_font As IFontDisp)
Set iFont = new_font
Label1.Font.Name = new_font.Name
Label1.Font.Size = new_font.Size
Label1.Font.Italic = new_font.Italic
Label1.Font.Underline = new_font.Underline
Label1.Font.Bold = new_font.Bold



PropertyChanged "Font"
End Property

Private Sub Image1_Click(Index As Integer)
RaiseEvent click
End Sub

Private Sub Image1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Picture1_MouseDown 0, 0, 0, 0
End Sub

Private Sub Image1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Picture1_MouseUp 0, 0, 0, 0
End Sub

Private Sub Label1_Click()
Call Image1_Click(0)
End Sub

Private Sub Label1_DblClick()
RaiseEvent click
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
 Picture1_MouseDown 0, 0, 0, 0
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Picture1_MouseMove 0, 0, 0, 0
End Sub

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Picture1_MouseUp 0, 0, 0, 0
End Sub

Private Sub Picture1_Click()
 Call Image1_Click(0)
End Sub

Private Sub Picture1_DblClick()
RaiseEvent dobleclick
End Sub



Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)


For i = Line1.LBound To Line1.UBound
    Line1(i).BorderColor = vbBlack
Next i
Picture1.BorderStyle = 1
hi 1, 0
'sonido = sndPlaySound(ByVal "c:Selected.WAV", 0)
RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)


hu True, 2
If iStyle = 1 Then
imagen 0, 2
End If
RaiseEvent MouseMove(Button, Shift, x, y)
If (iborderstyle = 1 And iappearance = 0) Then
Line1(4).Visible = True
Line1(5).Visible = True
Else
Line1(4).Visible = False
Line1(5).Visible = False
End If
colorin vbYellow
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)


For i = Line1.LBound To Line1.UBound
If i <= 1 Then
Line1(i).BorderColor = vbWhite
ElseIf i > 1 Then
Line1(i).BorderColor = &H808080
End If
Next i

If iborderstyle = 1 And iappearance = 0 Then
Picture1.BorderStyle = 1
Else
Picture1.BorderStyle = 0
End If
If ivalue = 1 Then
j = j + 1
DoEvents
hi 0, 1
End If
If iStyle = 1 Then
Label1.Left = Image1(0).Left + Image1(0).Width + 30
End If

RaiseEvent MouseUp(Button, Shift, x, y)
End Sub


Private Sub UserControl_Click()
RaiseEvent click
End Sub

Private Sub UserControl_Initialize()
Picture1.Move 0, 0
Caption = "Command"
BorderStyle = 2
Appearance = 1
UserControl.Width = 1200
Value = 0
Style = 0
BackColor = &H8000000F
ForeColor = vbBlack
Alingment = 2
Line1(4).Visible = False
Line1(5).Visible = False
Set iFont = Label1.Font
iFont.Name = "ms sans serif"
iFont.Size = "8"
iFont.Italic = False
iFont.Bold = True
Enabled = True
End Sub



Private Sub UserControl_Resize()
UserControl.Height = Picture1.Height
Picture1.Width = ScaleWidth - 10
Line1(2).X1 = Picture1.Width - 20
Line1(2).X2 = Picture1.Width - 20
Line1(1).X2 = Picture1.ScaleWidth - Line1(2).BorderWidth - 50
Line1(1).X2 = Picture1.ScaleWidth - Line1(2).BorderWidth - 50
Line1(3).X2 = Picture1.ScaleWidth - Line1(2).BorderWidth
Line1(3).X2 = Picture1.ScaleWidth
Line1(4).X2 = Picture1.ScaleWidth: Line1(4).X2 = Picture1.ScaleWidth
Line1(5).X1 = Picture1.Width - 50: Line1(5).X2 = Picture1.Width - 50
Shape1.Width = Picture1.ScaleWidth
End Sub
Private Sub Picture1_Resize()
If iStyle = 1 Then
Label1.Left = Image1(0).Left + Image1(0).Width + 100
ElseIf iStyle = 0 Then
Label1.Left = (Picture1.ScaleWidth / 2) - (Label1.Width - 470)
End If
End Sub



Public Property Get Picture() As StdPicture
Set Picture = ipicture
End Property
Public Property Set Picture(ByVal new_picture As StdPicture)
Set ipicture = new_picture
If iStyle = 1 Then
Set u = ImageList1(1).ListImages.Add(1, , new_picture)
Set Image1(2).Picture = new_picture

End If
PropertyChanged "Picture"
End Property
Public Property Get Caption() As String
Caption = icaption
End Property
Public Property Let Caption(ByVal new_caption As String)
icaption = new_caption
Label1.Caption = new_caption
PropertyChanged "Caption"
End Property
Public Property Get BorderStyle() As aborderstyle
BorderStyle = iborderstyle
End Property
Public Property Let BorderStyle(ByVal new_borderstyle As aborderstyle)
iborderstyle = new_borderstyle
If Not new_borderstyle = 2 Then
hu False, 2
Picture1.BorderStyle = new_borderstyle
Else
hu True, 2
Picture1.BorderStyle = 0
End If
PropertyChanged "BorderStyle"
End Property
Public Property Get Appearance() As aappearance
Appearance = iappearance
End Property
Public Property Let Appearance(ByVal new_appearance As aappearance)
iappearance = new_appearance
Picture1.Appearance = new_appearance
ibackcolor = vbWhite
PropertyChanged "Appearance"
End Property
Public Property Get Value() As avalue
Value = ivalue
End Property
Public Property Let Value(ByRef new_ivalue As avalue)
ivalue = new_ivalue
PropertyChanged "Value"
End Property
Public Property Get Style() As aStyle
Style = iStyle
End Property
Public Property Let Style(ByRef new_style As aStyle)
iStyle = new_style
Image1(0).Visible = True
If iStyle = 0 Then
Set Image1(0).Picture = LoadPicture("")
Image1(0).Visible = False
Else
Label1.Left = Image1(0).Left + Image1(0).Width + 50
End If
PropertyChanged "Style"
End Property
Public Property Get BackColor() As OLE_COLOR
BackColor = ibackcolor
End Property
Public Property Let BackColor(ByVal new_backcolor As OLE_COLOR)
ibackcolor = new_backcolor
Picture1.BackColor = new_backcolor
PropertyChanged "BackColor"
End Property
Public Property Get ForeColor() As OLE_COLOR
ForeColor = iforecolor
End Property
Public Property Let ForeColor(ByVal new_backcolor As OLE_COLOR)
iforecolor = new_backcolor
Label1.ForeColor = new_backcolor
PropertyChanged "ForeColor"
End Property
Public Property Get Alingment() As aAlingment
Alingment = ialingment
End Property
Public Property Let Alingment(ByRef new_alingment As aAlingment)
ialingment = new_alingment
Label1.Alignment = new_alingment
PropertyChanged "Alingment"
End Property
Public Property Get DisabledPicture() As StdPicture
Set DisabledPicture = Idisabledpicture
End Property
Public Property Set DisabledPicture(ByVal new_disabledpicture As StdPicture)
Set Idisabledpicture = new_disabledpicture
If iStyle = 1 Then
Set u = ImageList1(0).ListImages.Add(1, , new_disabledpicture)
Set Image1(1).Picture = new_disabledpicture
imagee 0, 1
End If
PropertyChanged "DisabledPicture"
End Property
Public Sub UserControl_ReadProperties(PropBag As PropertyBag)
With PropBag
 Font.Name = .ReadProperty("FontName", "ms sans serif")
 Font.Size = .ReadProperty("FontSize", "8")
 Font.Bold = .ReadProperty("FontBold", True)
 Font.Italic = .ReadProperty("FontItalic", False)
End With
For i = Image1.LBound To Image1.UBound - 1
 Set Image1(i).Picture = PropBag.ReadProperty("Disabledpicture", Nothing)
Next i
Set Image1(2).Picture = PropBag.ReadProperty("Picture", Nothing)
Caption = PropBag.ReadProperty("Caption", "Command")
BorderStyle = PropBag.ReadProperty("BorderStyle", 2)
Appearance = PropBag.ReadProperty("Appearance", 0)
Value = PropBag.ReadProperty("Value", 0)
Style = PropBag.ReadProperty("Style", 0)
BackColor = PropBag.ReadProperty("Backcolor", &H8000000F)
ForeColor = PropBag.ReadProperty("ForeColor", vbBlack)
Alingment = PropBag.ReadProperty("Alingment", 2)
Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

Public Sub UserControl_WriteProperties(PropBag As PropertyBag)
Call PropBag.WriteProperty("FontName", Font.Name, "ms sans serif")
Call PropBag.WriteProperty("FontSize", Font.Size, "8")
Call PropBag.WriteProperty("FontBold", Font.Bold, True)
Call PropBag.WriteProperty("FontItalic", Font.Italic, False)
Call PropBag.WriteProperty("Enabled", ienabled, True)
For i = Image1.LBound To Image1.UBound - 1
Call PropBag.WriteProperty("Disabledpicture", Image1(i).Picture, Nothing)
Next i
Call PropBag.WriteProperty("Picture", Image1(2).Picture, Nothing)
Call PropBag.WriteProperty("Caption", icaption, "Command")
Call PropBag.WriteProperty("BorderStyle", iborderstyle, 2)
Call PropBag.WriteProperty("Appearance", iappearance, 0)
Call PropBag.WriteProperty("Value", ivalue, 0)
Call PropBag.WriteProperty("Style", iStyle, 0)
Call PropBag.WriteProperty("BackColor", ibackcolor, &H8000000F)
Call PropBag.WriteProperty("ForeColor", iforecolor, vbBlack)
Call PropBag.WriteProperty("Alingment", ialingment, 2)
End Sub


Public Function hu(ByRef ajo As Boolean, al As Integer)
For i = Line1.LBound To Line1.UBound - al
Line1(i).Visible = ajo
Next i
End Function
Public Function hi(ByRef nu1, nu2 As Integer)
If ivalue = 1 Then
If j = 1 Then
Picture1.BorderStyle = nu1
ElseIf j = 2 Then
Picture1.BorderStyle = nu2
j = 0
End If
End If

End Function
Public Sub imagee(ByRef n1, n As Integer)
Set Image1(0).Picture = ImageList1(n1).ListImages.Item(n).Picture
If Image1(0).Height > Picture1.ScaleHeight Then
Image1(0).Height = Picture1.ScaleHeight - 10
End If
Image1(0).Visible = True
End Sub
Public Sub imagen(ByRef j1 As Integer, j2 As Integer)
Image1(j1).Picture = Image1(j2).Picture
End Sub
Public Sub med()
If iStyle = 1 Then
Label1.Left = Image1(0).Left + Image1(0).Width + 50
ElseIf iStyle = 0 Then
Label1.Left = (Picture1.ScaleWidth / 2) - (Label1.Width - 470)
End If

End Sub
Public Sub action()
For Each Control In FrmTela(x).Controls
    If Control.Tag <> "a" Then
        hu False, 0
        imagen 0, 1
    End If
Next Control
End Sub
Public Sub colorin(color As Variant)
If iStyle = 1 Then
ForeColor = color
End If
End Sub
