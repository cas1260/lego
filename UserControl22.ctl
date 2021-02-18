VERSION 5.00
Begin VB.UserControl UserControl2 
   ClientHeight    =   3732
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4008
   ControlContainer=   -1  'True
   ScaleHeight     =   3732
   ScaleWidth      =   4008
   ToolboxBitmap   =   "UserControl22.ctx":0000
   Begin VB.Image Image1 
      Height          =   132
      Index           =   11
      Left            =   3216
      Picture         =   "UserControl22.ctx":0312
      Top             =   432
      Width           =   132
   End
   Begin VB.Image Image1 
      Height          =   132
      Index           =   8
      Left            =   2880
      Picture         =   "UserControl22.ctx":0C9F
      Top             =   2880
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Image Image1 
      Height          =   132
      Index           =   10
      Left            =   3120
      Picture         =   "UserControl22.ctx":15C7
      Top             =   3240
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Image Image1 
      Height          =   132
      Index           =   9
      Left            =   2760
      Picture         =   "UserControl22.ctx":1F54
      Top             =   3240
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Form1"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   885
      TabIndex        =   0
      Top             =   405
      Width           =   435
   End
   Begin VB.Image Image1 
      Height          =   132
      Index           =   7
      Left            =   720
      Picture         =   "UserControl22.ctx":2911
      Top             =   3360
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Image Image1 
      Height          =   132
      Index           =   6
      Left            =   2160
      Picture         =   "UserControl22.ctx":32B5
      Top             =   3480
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Image Image1 
      Height          =   132
      Index           =   3
      Left            =   3420
      Picture         =   "UserControl22.ctx":3C59
      Top             =   432
      Width           =   132
   End
   Begin VB.Image Image1 
      Height          =   132
      Index           =   1
      Left            =   3048
      Picture         =   "UserControl22.ctx":45FD
      Top             =   432
      Width           =   132
   End
   Begin VB.Image Image2 
      Height          =   1410
      Index           =   4
      Left            =   720
      Picture         =   "UserControl22.ctx":4FA1
      Stretch         =   -1  'True
      Top             =   720
      Width           =   45
   End
   Begin VB.Image Image1 
      Height          =   132
      Index           =   0
      Left            =   1080
      Picture         =   "UserControl22.ctx":5036
      Top             =   3120
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Image Image1 
      Height          =   132
      Index           =   2
      Left            =   1080
      Picture         =   "UserControl22.ctx":5A09
      Top             =   3480
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Image Image2 
      Height          =   315
      Index           =   1
      Left            =   720
      Picture         =   "UserControl22.ctx":6336
      Stretch         =   -1  'True
      Top             =   360
      Width           =   1530
   End
   Begin VB.Image Image1 
      Height          =   132
      Index           =   5
      Left            =   2160
      Picture         =   "UserControl22.ctx":654A
      Top             =   3000
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Image Image2 
      Height          =   1530
      Index           =   3
      Left            =   3600
      Picture         =   "UserControl22.ctx":6EAD
      Stretch         =   -1  'True
      Top             =   720
      Width           =   45
   End
   Begin VB.Image Image2 
      Height          =   315
      Index           =   2
      Left            =   2160
      Picture         =   "UserControl22.ctx":6F42
      Stretch         =   -1  'True
      Top             =   360
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   132
      Index           =   4
      Left            =   2160
      Picture         =   "UserControl22.ctx":7145
      Top             =   3240
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Image Image2 
      Height          =   45
      Index           =   0
      Left            =   720
      Picture         =   "UserControl22.ctx":7B3B
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   2880
   End
End
Attribute VB_Name = "UserControl2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Event resize()
Event click()
Event dobleclick()
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Event exits(Index As Integer)
Dim m(3) As Long
Dim k As Integer
Dim icontrolbox As Boolean
Dim icaption As String
Dim iautosize As Boolean
Dim iforecolor As OLE_COLOR
Dim ibackcolor As OLE_COLOR
Dim iFont As IFontDisp
Dim ipicture As StdPicture
Public Property Get ControlBox() As Boolean
ControlBox = icontrolbox
End Property
Public Property Let ControlBox(ByVal new_controlbox As Boolean)
icontrolbox = new_controlbox
If Not new_controlbox Then
Image1(1).Visible = False
Image1(11).Visible = False
Else
Image1(1).Visible = True
Image1(11).Visible = True
End If

PropertyChanged "ControlBox"
End Property
Public Property Get ForeColor() As OLE_COLOR
ForeColor = iforecolor
End Property
Public Property Let ForeColor(ByVal new_forecolor As OLE_COLOR)
iforecolor = new_forecolor
Label1.ForeColor = new_forecolor
PropertyChanged "ForeColor"
End Property
Public Property Get BackColor() As OLE_COLOR
BackColor = ibackcolor
End Property
Public Property Let BackColor(ByVal new_backcolor As OLE_COLOR)
ibackcolor = new_backcolor
UserControl.BackColor = new_backcolor
PropertyChanged "BackColor"
End Property
Public Property Get Font() As IFontDisp
Set Font = iFont
End Property
Public Property Set Font(ByVal new_font As IFontDisp)
iFont = new_font
 Label1.Font.Name = iFont.Name
 Label1.Font.size = iFont.size
 Label1.Font.Bold = iFont.Bold
PropertyChanged "Font"
End Property

Private Sub Image1_Click(Index As Integer)


RaiseEvent exits(Index)

End Sub

Private Sub Image1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Select Case Index
Case 1
Image1(Index).Picture = Image1(2).Picture
Case 3
Image1(Index).Picture = Image1(5).Picture
Case 11
Image1(Index).Picture = Image1(8).Picture
End Select
End Sub

Private Sub Image1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Select Case Index
Case 1
Image1(Index).Picture = Image1(0).Picture
Image1(3).Picture = Image1(6).Picture
Image1(11).Picture = Image1(10).Picture
Case 3
Image1(Index).Picture = Image1(4).Picture
Image1(1).Picture = Image1(7).Picture
Image1(11).Picture = Image1(10).Picture
Case 11
Image1(Index).Picture = Image1(9).Picture
Image1(1).Picture = Image1(7).Picture
Image1(3).Picture = Image1(6).Picture
End Select
End Sub

Private Sub Image1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Select Case Index
Case 1
Image1(Index).Picture = Image1(7).Picture
Case 3
Image1(Index).Picture = Image1(6).Picture
Case 11
Image1(Index).Picture = Image1(10).Picture
End Select
End Sub

Private Sub Image2_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
mover UserControl.Parent
End Sub

Private Sub Image2_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Image1(11).Picture = Image1(10).Picture
Image1(3).Picture = Image1(6).Picture
Image1(1).Picture = Image1(7).Picture
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Image2_MouseDown 0, 0, 0, 0, 0
End Sub

Private Sub UserControl_Click()
RaiseEvent click
End Sub

Private Sub UserControl_DblClick()
RaiseEvent dobleclick
End Sub

Private Sub UserControl_Initialize()
Image2(1).Move 0, 0
Image2(2).Move Image2(1).Width, 0
Image2(4).Move 0, Image2(1).Height
Image2(0).Move 0, Image2(4).Height + Image2(1).Height
Image2(3).Move Image2(2).Left + Image2(2).Width - Image2(3).Width, Image2(2).Height
Image2(0).Width = Image2(1).Width + Image2(2).Width
Image2(3).Height = Image2(4).Height
Image1(11).Move Image2(2).Left + Image2(2).Width + 580, 50
Image1(3).Move Image2(2).Left + Image2(2).Width - 300, 50
Image1(1).Move Image1(11).Left + Image1(1).Width + 15, 50
Label1.Move 130, 50
Caption = "Form1"
AutoSize = False
ForeColor = vbWhite
BackColor = &HC0C0C0
ControlBox = True
Set iFont = Label1.Font
iFont.Name = "ms sans serif"
iFont.size = "8"
iFont.Bold = False
ControlBox = True
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Image1(3).Picture = Image1(6).Picture
Image1(1).Picture = Image1(7).Picture
Image1(11).Picture = Image1(10).Picture
RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

Private Sub UserControl_Resize()
On Error Resume Next
If AutoSize Then
UserControl.Height = UserControl.Parent.ScaleHeight
UserControl.Width = UserControl.Parent.ScaleWidth
End If
Image2(2).Width = UserControl.ScaleWidth - Image2(1).Width
Image2(3).Left = UserControl.ScaleWidth - Image2(3).Width: Image2(3).Height = UserControl.ScaleHeight - (Image2(2).Height + Image2(0).Height)
Image2(0).Width = UserControl.ScaleWidth: Image2(0).Top = UserControl.ScaleHeight - (Image2(0).Height)
Image2(4).Height = UserControl.ScaleHeight - (Image2(1).Height + Image2(0).Height)
Image1(3).Left = UserControl.ScaleWidth - (Image1(3).Width * 2)
Image1(11).Left = UserControl.ScaleWidth - ((Image1(1).Width * 3) + 30)
Image1(1).Left = UserControl.ScaleWidth - ((Image1(1).Width * 4) + 40)
RaiseEvent resize
End Sub
Public Property Get Caption() As String
Caption = icaption
End Property
Public Property Let Caption(ByVal new_caption As String)
icaption = new_caption
Label1.Caption = new_caption
PropertyChanged "Caption"
End Property
Public Property Get AutoSize() As Boolean
AutoSize = iautosize
End Property
Public Property Let AutoSize(ByVal new_autosize As Boolean)
iautosize = new_autosize
Call UserControl_Resize
PropertyChanged "Autosize"
End Property
Public Property Get Picture() As StdPicture
Set Picture = ipicture
End Property
Public Property Set Picture(ByVal new_picture As StdPicture)
Set ipicture = new_picture
Set UserControl.Picture = new_picture
PropertyChanged "Picture"
End Property
Public Sub mover(UserControl As Form)
ReleaseCapture

SendMessage UserControl.hwnd, &HA1, 2, 0&
End Sub
Public Sub UserControl_ReadProperties(PropBag As PropertyBag)
With PropBag
AutoSize = .ReadProperty("AutoSize", False)
Caption = .ReadProperty("Caption", "Form1")
Set Picture = .ReadProperty("Picture", UserControl.Picture)
ForeColor = .ReadProperty("ForeColor", vbWhite)
BackColor = .ReadProperty("backcolor", &H8000000F)
Font.Name = .ReadProperty("FontName", "ms sans serif")
Font.size = .ReadProperty("FontSize", "8")
Font.Bold = .ReadProperty("FontBould", False)
ControlBox = .ReadProperty("ControlBox", True)
End With
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
Call PropBag.WriteProperty("AutoSize", iautosize, False)
Call PropBag.WriteProperty("Caption", icaption, "Form1")
Call PropBag.WriteProperty("Picture", UserControl.Picture)
Call PropBag.WriteProperty("ForeColor", iforecolor, vbWhite)
Call PropBag.WriteProperty("FontName", Font.Name, "ms sans serif")
Call PropBag.WriteProperty("FontSize", Font.size, "8")
Call PropBag.WriteProperty("FontBold", Font.Bold, False)
Call PropBag.WriteProperty("ControlBox", icontrolbox, True)
End Sub
Public Sub restaurar()
k = k + 1
If k = 1 Then
m(0) = Form1.Left
m(1) = Form1.Top
m(2) = Form1.Width
m(3) = Form1.Height
Form1.Move 0, 0
Form1.Width = Screen.Width
Form1.Height = Screen.Height - 400
UserControl.Width = UserControl.Parent.Width
UserControl.Height = UserControl.Parent.Height
ElseIf k = 2 Then

Form1.Left = m(0)
Form1.Top = m(1)
Form1.Width = m(2)
Form1.Height = m(3)
Width = m(2)
Height = m(3)
k = 0

End If
End Sub
