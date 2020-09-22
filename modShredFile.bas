Attribute VB_Name = "modShredFile"
'Please refer to HELP (modHelp.bas) for documentation.
'-----------------------------------------------------

Global NumberOfTimes As Long 'number of times we should
'overwrite BINARY, configurable by the user

Global Setting As String 'ultra quick, paranoid etc

Global Method As String 'DOD or shredfile?

Global Rename As Boolean 'rename the file?

Global FileTemp As String

Declare Function FlushFileBuffers Lib "kernel32" (ByVal hfile As Long) As Long
'flush file buffers - DO NOT EVER REMOVE!

Public Sub ShredFile(sFileName As String)

    On Error GoTo ErrSub
    
    If Method = "Dod" Then
    Exit Sub
    DodWipeFile sFileName, 9
    ElseIf Method = "ShredFile" Then
    End If
    
    Dim i, x As Long
    
    '=============================
    
    Open sFileName For Binary As #1
    
    For x = 1 To NumberOfTimes 'loop until satisfied
    
    frmMain.PB1.Value = 0: frmMain.PB1.Max = NumberOfTimes: frmMain.PB1.Value = x
    frmMain.SB1.Panels(1).Text = "Overwriting " & "... " & x & " of " & NumberOfTimes
    
    For i = 1 To LOF(1)
    Put #1, i, RandomBin(1)
    FlushFileBuffers (1)
    
    Next i
    Next x
    
    Close #1
    
    '=============================
    
    If Rename = True Then
    
    Dim RandomName, OldName, NewName As String
    Dim a, b As Long
    
    a = Rnd * 1000 'a & b are random numbers which we
    b = Rnd * 1000 'use to generate a new file name
    
    RandomName = a & "." & b 'the final random name
    'in the format of 111.111.111

    OldName = FileTemp
    NewName = GetPath(FileTemp) & RandomName

    Name OldName As NewName

    End If

    Kill NewName 'kill it

ErrSub: 'should an error occur

    If Err.Number = 0 Then

    If Err.Number = 9 Then 'probably hex corrupt generated a figure of ""
    Resume Next

    If Err.Number = 55 Then
    Close #1 'if file is already open, close it
    
    Else
    MsgBox (Err.Number & vbCrLf & Err.Description), vbCritical + vbOKOnly, "Error"

    End If: End If: End If
    

End Sub

