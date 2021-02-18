Attribute VB_Name = "FileOperation"
Public Function SaveFile1(FileName As String, Text As String)
    If Dir(FileName) <> "" Then
        Kill FileName
    End If
    Open FileName For Binary Access Write As #1
    
        If Err Then
'            MsgBox "Impossivel Salvar, Error na Leitura do 0Arquivo  ", vbCritical, App.Title
            Exit Function
        End If
        
    Put #1, , Text
    Close #1

End Function
Public Function SaveFile(FileName, Text As String)

    Open FileName For Output As #1
        If Err Then
           ' MsgBox "Impossivel Salvar, Error na Leitura do 0Arquivo  ", vbCritical, App.Title
            Exit Function
        End If
        'Put #1, , Text
        Print #1, , Text
    Close #1

End Function
Function OpenFile1(FileName) As String
On Error GoTo Trata_Erro
    If Dir(FileName) = "" Then
        'Resp 17, ""
        Exit Function
    End If
    Open FileName For Input As #1
        If Err Then
            'MsgBox "Impossivel Abrir, Error na Leitura do Arquivo ", vbExclamation, App.Title
            Exit Function
        End If
        OpenFile1 = Input(LOF(1), 1)
    Close #1
Trata_Erro:

End Function
