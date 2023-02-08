Attribute VB_Name = "CompressionService"
Option Explicit

Private Const FOF_ALLOWUNDO = &H40
Private Const FOF_NOCONFIRMATION = &H10
Private Const FO_COPY = &H2

Sub UnzipFile(sourceFile As String, destinationPath As String)

Dim FSO As FileSystemObject: Set FSO = New FileSystemObject

Dim thisFile As Scripting.File
Dim originalName As String

Dim oShell As Object
Dim fileSource As Object
Dim fileDest As Object

    On Error GoTo EH

    Set thisFile = FSO.GetFile(sourceFile)
    originalName = thisFile.Name
    
    If LCase$(Right$(thisFile.Name, 4)) <> ".zip" Then
        originalName = thisFile.Name
        thisFile.Name = thisFile.Name & ".zip"
    End If

    If sourceFile = "" Or destinationPath = "" Then GoTo EH
    
    Set oShell = CreateObject("Shell.Application")
    If oShell Is Nothing Then GoTo EH
 
    If Right$(UCase$(sourceFile), 4) <> ".ZIP" Then
        sourceFile = sourceFile & ".ZIP"
    End If
    
    Set fileSource = oShell.NameSpace("" & sourceFile)      '//should be zip file
    Set fileDest = oShell.NameSpace("" & destinationPath)          '//should be directory

    Call fileDest.CopyHere(fileSource.Items, FOF_NOCONFIRMATION)
    If thisFile.Name <> originalName Then thisFile.Name = originalName

EH:
    Set oShell = Nothing
    Set fileSource = Nothing
    Set fileDest = Nothing
    Exit Sub

    If Err.Number = 70 Then
        MsgBox "File access error", vbCritical
    Else
        MsgBox "There was a problem installing the app." & vbCrLf & "Makesure the app can write to " & destinationPath, vbExclamation, "error"
        
        On Error Resume Next
        If originalName <> vbNullString Then
            If thisFile.Name <> originalName Then thisFile.Name = originalName
        End If
    End If
End Sub
