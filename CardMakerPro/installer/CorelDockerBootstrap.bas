Attribute VB_Name = "CardMakerProBootstrap"
Option Explicit

' Import into CorelDRAW VBA editor as a standard module.
' Set dllPath to the extracted GitHub Actions host DLL location.
' Uses CorelDRAW's documented FrameWork.AddDocker API.
Public Sub RegisterCardMakerProDocker()
    Dim dllPath As String
    dllPath = InputBox("Full path to CardMakerPro.Host.dll:", "Card Maker Pro - Docker Registration")
    If Len(Trim$(dllPath)) = 0 Then Exit Sub
    If Dir$(dllPath) = "" Then
        MsgBox "Host DLL not found: " & dllPath, vbExclamation
        Exit Sub
    End If
    On Error GoTo RegistrationError
    Application.FrameWork.AddDocker "F46B2D41-831D-47C8-95C5-B4CC7C0A01A2", _
        "CardMakerPro.Host.DockerControl", dllPath
    MsgBox "CorelDRAW accepted the AddDocker request." & vbCrLf & _
           "Check Window > Dockers. Actual visibility and WPF loading still require testing.", vbInformation
    Exit Sub
RegistrationError:
    MsgBox "Docker registration failed (" & Err.Number & "): " & Err.Description & vbCrLf & _
           "Verify DLL architecture, dependencies, CorelDRAW edition and .NET Framework.", vbCritical
End Sub
