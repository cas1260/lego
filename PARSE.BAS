Attribute VB_Name = "Module1"
Option Explicit

Function Parse$(ByVal sString As String, iReq As Integer, sDelim As String, Optional Padrao As String)
Dim sSt As String, iCnt As Integer, iPos As Integer

    If Len(sDelim) = 0 Then sDelim = ","
    sSt = sString & sDelim
    For iCnt = 1 To iReq
        iPos = InStr(sSt, sDelim)
        If iPos Then
            If iCnt = iReq Then     ' Requested string
                Parse$ = Left$(sSt, iPos - 1)
                Exit For
            End If
            If iPos = Len(sSt) Then ' No string left
                Parse$ = Padrao
                Exit For
            End If
            sSt = Mid$(sSt, iPos + Len(sDelim))
        Else
            Parse$ = sSt
            Exit For
        End If
    Next iCnt
End Function

