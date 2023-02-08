Attribute VB_Name = "MainService"
Option Explicit

Private Const PRODUCT_NAME As String = "ViPad"

Private Const APP_EXE As String = "ViPad.exe"
Private Const APP_ACTIVATOR_EXE As String = "ViPadActivator.exe"

Sub Main()
    On Error GoTo hErr

Dim installPath As String: installPath = Environ$("AppData") & "\" & PRODUCT_NAME
Dim appArchieve As String: appArchieve = installPath & "\appdata.zip"
Dim appExePath As String: appExePath = installPath & "\" & APP_EXE
Dim appActivatorPath As String: appActivatorPath = installPath & "\" & APP_ACTIVATOR_EXE

Dim FSO As FileSystemObject: Set FSO = New FileSystemObject

    If FSO.FolderExists(installPath) = False Then
        FSO.CreateFolder installPath
    End If

    DumpAppData appArchieve
    
    CompressionService.UnzipFile appArchieve, installPath
    
    If FSO.FileExists(appActivatorPath) Then
        Shell """" & appActivatorPath & """" & " /auto"
    End If
    Shell appExePath
    
    FSO.DeleteFile appArchieve
    FSO.DeleteFile appActivatorPath
    
    Exit Sub
hErr:
    MsgBox "An error occured during the extraction", vbCritical
End Sub

Sub DumpAppData(ByVal filePath As String)

On Error GoTo Handler

    Dim byteData() As Byte
    Dim i As Long

    byteData = LoadResData("APPDATA", "Custom")
    Open filePath For Binary As #1
    
    For i = LBound(byteData) To UBound(byteData)
        Put #1, , byteData(i)
    Next i
    
    Close #1
    
    Exit Sub
Handler:
    MsgBox Err.Description, vbCritical
End Sub
