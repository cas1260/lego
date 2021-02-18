VERSION 5.00
Begin VB.UserControl UserControl3 
   ClientHeight    =   2640
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4920
   ScaleHeight     =   2640
   ScaleWidth      =   4920
   ToolboxBitmap   =   "UserControl33.ctx":0000
   Begin VB.PictureBox Picture1 
      Height          =   375
      Left            =   600
      ScaleHeight     =   324
      ScaleWidth      =   3444
      TabIndex        =   0
      Top             =   960
      Width           =   3495
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0.00%"
         Height          =   195
         Left            =   1560
         TabIndex        =   2
         Top             =   60
         Width           =   435
      End
      Begin VB.Label Label1 
         Height          =   255
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   15
      End
   End
End
Attribute VB_Name = "UserControl3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Event click()
Event dobleclick()

Public Enum eborderstyle
None = 0
Fixed = 1
End Enum


Public Enum eappearance
Flat = 0
cc3D = 1
End Enum

Dim scala As Long
Dim ibackcolor As OLE_COLOR
Dim ifillcolor As OLE_COLOR
Dim iforecolor As OLE_COLOR
Dim iborderstyle As eborderstyle
Dim iappearance As aappearance
Dim imin, imax, ivalue As Long
Public Property Get BackColor() As OLE_COLOR
BackColor = ibackcolor
End Property
Public Property Let BackColor(ByVal new_backcolor As OLE_COLOR)
ibackcolor = new_backcolor
Label1.BackColor = new_backcolor
PropertyChanged "BackColor"
End Property
Public Property Get ForeColor() As OLE_COLOR
ForeColor = iforecolor
End Property
Public Property Let ForeColor(ByVal new_forecolor As OLE_COLOR)
iforecolor = new_forecolor
Label2.ForeColor = new_forecolor
PropertyChanged "BackColor"
End Property
Public Property Get FillColor() As OLE_COLOR
FillColor = ifillcolor
End Property
Public Property Let FillColor(ByVal new_fillcolor As OLE_COLOR)
ifillcolor = new_fillcolor
Picture1.BackColor = new_fillcolor
PropertyChanged "FillColor"
End Property
Public Property Get BorderStyle() As eborderstyle
BorderStyle = iborderstyle
End Property
Public Property Let BorderStyle(ByVal new_borderstyle As eborderstyle)
iborderstyle = new_borderstyle
Picture1.BorderStyle = new_borderstyle
PropertyChanged "BorderStyle"
End Property
Public Property Get Appearance() As eappearance
Appearance = iappearance
End Property
Public Property Let Appearance(ByVal new_appearance As eappearance)
iappearance = new_appearance
Picture1.Appearance = new_appearance
If new_appearance = 0 Then
ifillcolor = vbWhite
Call UserControl_Resize
End If
PropertyChanged "Apperance"
End Property

Public Property Get Min() As Long
Min = imin
End Property
Public Property Let Min(ByVal new_min As Long)
imin = new_min
If imin > imax Then
imin = imax
ElseIf imin > ivalue Then
ivalue = imin
End If
PropertyChanged "Min"
End Property
Public Property Get Max() As Long
Max = imax
End Property
Public Property Let Max(ByVal new_max As Long)
imax = new_max
If imax < imin Then
imax = imin
ElseIf ivalue > imax Then
imax = ivalue
End If
PropertyChanged "Max"
End Property
Public Property Get Value() As Long
Value = ivalue
End Property
Public Property Let Value(ByVal new_value As Long)
ivalue = new_value
If ivalue > imax Then
ivalue = imax
ElseIf ivalue < imin Then
ivalue = imin
End If
Label1.Visible = True
scala = (ivalue - imin) / (imax - imin) * Picture1.ScaleWidth
Label2.Caption = Format(ivalue / imax, "0.00%")
Label1.Width = scala

PropertyChanged "Value"
End Property


Private Sub Label1_Click()
Call Picture1_Click
End Sub

Private Sub Label1_DblClick()
RaiseEvent dobleclick
End Sub

Private Sub Picture1_Click()
RaiseEvent click
End Sub

Private Sub Picture1_DblClick()
RaiseEvent dobleclick
End Sub

Private Sub UserControl_Initialize()
Picture1.Move 0, 0
Label1.Height = Picture1.ScaleHeight
FillColor = &H8000000F
BackColor = vbRed
BorderStyle = 1
Label1.Visible = False
Appearance = 1
Min = 0
Max = 100
Value = 1
ForeColor = vbBlack
End Sub
Private Sub UserControl_Resize()
Picture1.Width = ScaleWidth
Picture1.Height = ScaleHeight
Label1.Height = Picture1.ScaleHeight
Label2.Left = (Picture1.ScaleWidth / 2) - (Label2.Width / 2)
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
FillColor = PropBag.ReadProperty("FillColor", &H8000000F)
BackColor = PropBag.ReadProperty("BackColor", vbRed)
BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
Appearance = PropBag.ReadProperty("Appearance", 1)
Min = PropBag.ReadProperty("Min", 0)
Max = PropBag.ReadProperty("Max", 100)
Value = PropBag.ReadProperty("Value", 1)
ForeColor = PropBag.ReadProperty("ForeColor", vbBlack)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
Call PropBag.WriteProperty("FillColor", ifillcolor, &H8000000F)
Call PropBag.WriteProperty("BackColor", ibackcolor, vbRed)
Call PropBag.WriteProperty("BorderStyle", iborderstyle, 1)
Call PropBag.WriteProperty("Appearance", iappearance, 1)
Call PropBag.WriteProperty("Min", imin, 0)
Call PropBag.WriteProperty("Max", imax, 100)
Call PropBag.WriteProperty("Value", ivalue, 1)
Call PropBag.WriteProperty("ForeColor", iforecolor, vbBlack)
End Sub
