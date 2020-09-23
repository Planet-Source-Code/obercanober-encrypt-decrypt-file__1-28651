Attribute VB_Name = "FileConverter"
Private Const CPY_BUFFER = 10240
Private Const DRY_FILE = ":\Temp.$$$"

Public Function ConvertFile(ByVal sSource As String, ByVal sDestination As String) As Boolean
    Dim SNum As Integer
    Dim DNum As Integer
    Dim perc As Integer
    Dim DFC As Double
    Dim DFCC As Double
    Dim ST As String
    
    If Dir(sSource) = Empty Then GoTo ERROR_HANDLER


    SNum = FreeFile
    DNum = FreeFile + 1
    Open sSource For Binary As SNum
        If LOF(SNum) = 0 Then
            Close SNum
            GoTo ERROR_HANDLER
        End If
        
        Open sDestination For Binary As DNum
            If LOF(SNum) > CPY_BUFFER Then
                ST = Space(CPY_BUFFER)
                DFC = Int(LOF(SNum) / CPY_BUFFER)
                DFCC = LOF(SNum)
                For i = 1 To DFC
                    DoEvents
                    Get SNum, , ST
                    Put DNum, , Convert(ST)
                    DFCC = DFCC - CPY_BUFFER
                    perc = Int((LOF(DNum) / LOF(SNum) * 100) + 1)
                Next i
                ST = String(DFCC, CPY_SPACE)
                Get SNum, , ST
                Put DNum, , ST
            ElseIf LOF(SNum) <= CPY_BUFFER And LOF(SNum) > 0 Then
                ST = Space(LOF(SNum))
                Get SNum, , ST
                Put DNum, , Convert(ST)
                perc = Int((LOF(DNum) / LOF(SNum) * 100) + 1)
            End If
            ST = Empty
            
        Close SNum
    Close DNum
        Exit Function

ERROR_HANDLER:
    Close SNum
    Close DNum
End Function

Private Function FileExist(ByVal sFile As String) As Boolean
    FileExist = True
    If Dir(sFile) = Empty Then FileExist = False
End Function

Private Function Convert(cString As String) As String
  
    For cCode = 1 To Len(cString)
        Convert = Convert + Chr(255 - Asc(Mid(cString, CInt(cCode), 1)))
    Next cCode

End Function


