Attribute VB_Name = "mPublic"
Option Explicit

' ***********************************************************************
' ShellExecute hwnd, "Open", sLink, "", App.path, 1
' ***********************************************************************
Declare Function ShellExecute Lib "Shell32.dll" Alias "ShellExecuteA" _
        (ByVal hWnd As Long, _
        ByVal lpOperation As String, _
        ByVal lpFile As String, _
        ByVal lpParameters As String, _
        ByVal lpDirectory As String, _
        ByVal nShowCmd As Long) As Long


Public Sub ShowError(Optional ByVal Module As String = "")
    
    MsgBox "Err:" & Err.Number & " " & Err.Description, vbCritical, "Error:" & Trim(Module)
    
End Sub

Public Function FileLines(sFile As String) As Long
Dim iFile           As Integer
Dim lCount          As Long
Dim sstr            As String

    If Dir(sFile) = "" Then
        FileLines = 0
        MsgBox "file not found", vbInformation
        Exit Function
    End If
    
    iFile = FreeFile
    lCount = 0
    
    Open sFile For Input As #iFile
    
    While Not EOF(iFile)
        Line Input #iFile, sstr
        If Trim(sstr) <> "" Then
            lCount = lCount + 1
        End If
    Wend
    
    Close #iFile
    
    FileLines = lCount
    

End Function

Public Function GetFileName(CD As Object, _
                            Title As String, _
                            FileTypes As String, _
                            DefType As String, _
                            Optional OpenDlg As Boolean = True) As String
On Error GoTo errhandler
    CD.CancelError = True
    CD.FLAGS = cdlCFEffects Or cdlCFBoth
    CD.Filter = FileTypes
'    CD.Filename = DefType
    CD.DialogTitle = Title
    If OpenDlg = True Then
        CD.ShowOpen
    Else
        CD.ShowSave
    End If

    GetFileName = CD.Filename

    Exit Function
errhandler:
    ShowError
    GetFileName = ""

End Function

Public Function Filename(ByVal Path As String, Optional IncExt As Boolean = True) As String
Dim l, i, P, lp As Integer
Dim rString As String
    rString = ""
    P = InStr(1, Path, "\")
    lp = P
    While P <> 0
        lp = P + 1
        P = InStr(lp, Path, "\")
    Wend
    If lp = 0 Then
        rString = Path
    Else
        rString = Mid(Path, lp, Len(Path) - lp + 1)
    End If

    If IncExt = False And InStr(1, rString, ".") > 0 Then
        rString = Mid(rString, 1, InStr(1, rString, ".") - 1)
    End If

    Filename = rString
End Function


Public Function FilePath(ByVal Path As String, Optional IncSlash As Boolean = True) As String
Dim l, i, P, lp As Integer
Dim rString As String
    rString = ""
    P = InStr(1, Path, "\")
    lp = P
    While P <> 0
        lp = P + 1
        P = InStr(lp, Path, "\")
    Wend
    
    If lp = 0 Then
        rString = ""
    Else
        If IncSlash Then
            rString = Left(Path, lp - 1)
        Else
            rString = Left(Path, lp - 2)
        End If
            
    End If


    FilePath = rString
End Function

