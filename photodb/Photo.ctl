VERSION 5.00
Begin VB.UserControl Photo 
   BackStyle       =   0  'Transparent
   ClientHeight    =   1275
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1200
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   1275
   ScaleWidth      =   1200
   ToolboxBitmap   =   "Photo.ctx":0000
   Begin VB.Image Def 
      Height          =   240
      Left            =   855
      Picture         =   "Photo.ctx":0312
      Top             =   855
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Photo 
      Height          =   1185
      Left            =   45
      OLEDropMode     =   1  'Manual
      Stretch         =   -1  'True
      Top             =   45
      Width           =   1095
   End
   Begin VB.Shape PhotoFrame 
      Height          =   1185
      Left            =   45
      Top             =   45
      Width           =   1095
   End
End
Attribute VB_Name = "Photo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------
'Photo OCX Ver 1.0
'Support ADO and DAO.
'
'Has only be tested on an access database.!!!!!!!
'
'Disigned by Rodney Safe Computing Tiger software
'You are free to distribute this code.
'But do not forget to include my name somewhere in
'your comments.
'Have a nice.
'Rodney Godfried.
'------------------------------------------------------------------

Option Explicit
Enum Connect
    useAdo = 1
    useDao = 2
End Enum


Dim XFig As Long
Dim DataFile As Integer, FileLength As Long, Chunks As Integer
Dim SmallChunks As Integer, Chunk() As Byte, I As Integer
Const ChunkSize As Integer = 1024
Public PhotoFileName As String
Public Event OnPhotoSaving(Succeded As Boolean, FileName As String)
Public Event OnPhotoLoading(IsPicture As Boolean, ErrorDescription As String)
Const m_def_ConnectionType = 1
Dim m_ConnectionType As Connect

Public Sub Reset()
    '---------------------------------------------
    'Clear the Photo picture box
    '---------------------------------------------
    Photo.Picture = LoadPicture("")
End Sub
Public Sub Picture(Pic)
    If Pic <> 0 Then
        Photo.Picture = Pic
        SavePicture Pic, "C:\T" + Trim(Str(XFig)) + ".mar"
        PhotoFileName = "C:\T" + Trim(Str(XFig)) + ".mar"
        XFig = XFig + 1
    End If
End Sub
Public Sub Refresh()
    '---------------------------------------------
    'Load the current imagefile into the picture box
    '---------------------------------------------
    If Len(PhotoFileName) > 0 Then Photo.Picture = LoadPicture(PhotoFileName)
End Sub

Public Function OpenPhotoFile() As String
Dim Filter As String
Dim FileName As String
On Error GoTo Out
    '---------------------------------------------
    'Open a common dialog whitout ocx to browse
    'for an image file
    '---------------------------------------------

    Filter = "Pictures(*.bmp;*.ico;*.gif;*.jpg)|*.bmp;*.ico;*.gif;*.jpg|All Files (*.*)|*.*"
    PhotoFileName = OpenFile(Filter, "Select Photo Image", App.Path)
    OpenPhotoFile = PhotoFileName
    Photo.Picture = LoadPicture(PhotoFileName)
Exit Function
Out:
    MsgBox Err.Description
End Function

Public Sub SavePhoto(Fieldname As Field)
Dim RS As Recordset
On Error GoTo Out

'---------------------------------------------
' If there is no image file exits
'---------------------------------------------
If Len(PhotoFileName) = 0 Then Exit Sub
DataFile = 1

'---------------------------------------------
'Open the image file
'---------------------------------------------
Open PhotoFileName For Binary Access Read As DataFile
    FileLength = LOF(DataFile)    ' Length of data in file
    '---------------------------------------------
    'If the imagefile is empty exits
    '---------------------------------------------
    If FileLength = 0 Then
        Close DataFile
        Exit Sub
    End If
    '---------------------------------------------
    'Calculate the bytes(Chunks)pakages to write
    '---------------------------------------------
    Chunks = FileLength \ ChunkSize
    SmallChunks = FileLength Mod ChunkSize
    '---------------------------------------------
    'Resize the chunck array to adjust the firts bytes package
    'To be copied
    '---------------------------------------------
    
    ReDim Chunk(SmallChunks)
    Get DataFile, , Chunk()
    '---------------------------------------------
    'Write the bytes to the given database fieldname
    '---------------------------------------------
    Fieldname.AppendChunk Chunk()
    '---------------------------------------------
    'Adjust the chunck array for the rest bytes
    'packages to be copied
    '---------------------------------------------
    ReDim Chunk(ChunkSize)
    If Chunks = 0 Then
        Get DataFile, , Chunk()
        Fieldname.AppendChunk Chunk()
    Else
        For I = 1 To Chunks
            Get DataFile, , Chunk()
            Fieldname.AppendChunk Chunk()
        Next I
    End If
Close DataFile
RaiseEvent OnPhotoSaving(True, PhotoFileName)
Exit Sub
Out:
RaiseEvent OnPhotoSaving(False, PhotoFileName)
End Sub


Public Function LoadPhoto(Fieldname As Field) As String

Dim lngOffset As Long
Dim lngTotalSize As Long
Dim strChunk As String


On Error GoTo Out

DataFile = 1

Open App.Path & "\RscPic.tmp" For Binary Access Write As DataFile
   '============================================
   'Support ado and Dao
   'Choose the right command according to user connection type
   '============================================
   If m_ConnectionType = useAdo Then
        lngTotalSize = Fieldname.ActualSize
    Else
        lngTotalSize = DaoFieldSize(Fieldname)
    End If
    
    Chunks = lngTotalSize \ ChunkSize
    SmallChunks = lngTotalSize Mod ChunkSize
        
        ReDim Chunk(ChunkSize)
            '============================================
            'Support ado and Dao
            'Choose the right command according to user connection type
            '============================================
            
        If m_ConnectionType = useDao Then
            Chunk() = GetDaoChunk(Fieldname)
        Else
            Chunk() = Fieldname.GetChunk(ChunkSize)
        End If
        
        Put DataFile, , Chunk()
        lngOffset = lngOffset + ChunkSize
        
        Do While lngOffset < lngTotalSize
            '============================================
            'Support ado and Dao
            'Choose the right command according to user connection type
            '============================================
            
            If m_ConnectionType = useAdo Then
                 Chunk() = Fieldname.GetChunk(ChunkSize)
            Else
                 Chunk() = GetDaoChunk(Fieldname)
            End If
            Put DataFile, , Chunk()
            lngOffset = lngOffset + ChunkSize
        Loop
Close DataFile
'============================================
' Pass the image file location to the function
'============================================
LoadPhoto = App.Path & "\RscPic.tmp"

'============================================
'Load the picture into the image box
'============================================

Photo.Picture = LoadPicture(App.Path & "\RscPic.tmp")
RaiseEvent OnPhotoLoading(True, "")

Exit Function

Out:
Photo.Picture = Def.Picture
Err.Clear
RaiseEvent OnPhotoLoading(False, Err.Description)

End Function

'The fallowing function retrieve the fieldsize when
'Using a dao connection
Private Function DaoFieldSize(Fieldname As DAO.Field) As Long
Dim lngOffset As Long
    DaoFieldSize = Fieldname.FieldSize
End Function

'The fallowing function retrieve the Chunk data when
'Using a dao connection
Private Function GetDaoChunk(Fieldname As DAO.Field)
Dim lngOffset As Long
    GetDaoChunk = Fieldname.GetChunk(lngOffset, ChunkSize)
End Function

'The fallowing Sub  set the frame and resize it
'To the user size
Private Sub UserControl_Resize()
Photo.Move 20, 20, UserControl.Width - 20, UserControl.Height - 20
PhotoFrame.Move 10, 10, UserControl.Width - 10, UserControl.Height - 10
sHwnd = UserControl.hwnd
End Sub

Private Sub UserControl_InitProperties()
    m_ConnectionType = m_def_ConnectionType
    XFig = 0
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    PhotoFrame.BorderWidth = PropBag.ReadProperty("BorderWidth", 1)
    PhotoFrame.BorderColor = PropBag.ReadProperty("BorderColor", &H0&)
    PhotoFrame.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
    PhotoFrame.BackStyle = PropBag.ReadProperty("BackStyle", 0)
    PhotoFrame.BorderColor = PropBag.ReadProperty("BorderColor", &H80000008)
    m_ConnectionType = PropBag.ReadProperty("ConnectionType", m_def_ConnectionType)
    Photo.Stretch = PropBag.ReadProperty("Stretch", True)
End Sub

Private Sub UserControl_Terminate()
On Error Resume Next
Kill "C:\T*.mar"
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BorderColor", PhotoFrame.BorderColor, &H80000008)
    Call PropBag.WriteProperty("BorderWidth", PhotoFrame.BorderWidth, 1)
    Call PropBag.WriteProperty("BorderColor", PhotoFrame.BorderColor, &H0&)
    Call PropBag.WriteProperty("BorderWidth", PhotoFrame.BorderWidth, 1)
    Call PropBag.WriteProperty("BackStyle", PhotoFrame.BackStyle, 1)
    Call PropBag.WriteProperty("Stretch", Photo.Stretch, True)
    Call PropBag.WriteProperty("ConnectionType", m_ConnectionType, m_def_ConnectionType)
End Sub

Public Property Get ConnectionType() As Connect
Attribute ConnectionType.VB_Description = "Return which connection type is used ADO or Dao"
    ConnectionType = m_ConnectionType
End Property

Public Property Let ConnectionType(ByVal New_ConnectionType As Connect)
    m_ConnectionType = New_ConnectionType
    PropertyChanged "ConnectionType"
End Property

Public Property Get hwnd() As Long
    hwnd = UserControl.hwnd
End Property

Public Property Get Stretch() As Boolean
Attribute Stretch.VB_Description = "Returns/sets a value that determines whether a graphic resizes to fit the size of an Image control."
    Stretch = Photo.Stretch
End Property

Public Property Let Stretch(ByVal New_Stretch As Boolean)
    Photo.Stretch() = New_Stretch
    PropertyChanged "Stretch"
End Property

Public Property Get BorderColor() As OLE_COLOR
Attribute BorderColor.VB_Description = "Returns/sets the border style for an object."
    BorderColor = PhotoFrame.BorderColor
End Property

Public Property Let BorderColor(ByVal New_BorderColor As OLE_COLOR)
    PhotoFrame.BorderColor() = New_BorderColor
    PropertyChanged "BorderColor"
End Property

Public Property Get BorderWidth() As Integer
Attribute BorderWidth.VB_Description = "Returns or sets the width of a control's border."
    BorderWidth = PhotoFrame.BorderWidth
End Property

Public Property Let BorderWidth(ByVal New_BorderWidth As Integer)
    PhotoFrame.BorderWidth() = New_BorderWidth
    PropertyChanged "BorderWidth"
End Property


