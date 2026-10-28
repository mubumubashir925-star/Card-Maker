Attribute VB_Name = "CardMakerDesignerSetup"
Option Explicit

' Run once from VBA editor with a BACKUP of GlobalMacros.gms.
' This changes actual designer controls, not merely runtime styling.
' Requires "Trust access to the VBA project object model" and an unlocked project.
Public Sub ApplyModernCardMakerDesigner()
    Dim prj As Object, comp As Object, d As Object, c As Object
    Dim mainFrame As Object, batch As Object
    Dim iconNames As String, n As String, topY As Single, leftX As Single
    Dim baseW As Single, rightX As Single, i As Long, countMoved As Long
    Dim fr As Object
    On Error GoTo Failed
    Set prj = Application.VBE.ActiveVBProject
    Set comp = prj.VBComponents("Card_Maker")
    Set d = comp.Designer
    Set mainFrame = FindFrameByCaption(d, "Card Maker")
    If mainFrame Is Nothing Then Err.Raise 5, , "Card Maker frame not found."
    iconNames = "|FitImgtoPCframe|Stretch_toFit|FitTextToFrame|"
    mainFrame.Caption = "CARD DATA AND FIELD MAPPING"
    mainFrame.BackColor = RGB(255, 255, 255)
    mainFrame.ForeColor = RGB(30, 47, 74)
    mainFrame.Font.Name = "Segoe UI"
    mainFrame.Font.Size = 9
    baseW = mainFrame.Width
    rightX = baseW + 15
    mainFrame.Width = baseW + 245
    For Each c In mainFrame.Controls
        On Error Resume Next
        c.Font.Name = "Segoe UI"
        c.Font.Size = 9
        If TypeName(c) = "CommandButton" Then
            c.BackColor = RGB(227, 237, 248)
            c.ForeColor = RGB(30, 47, 74)
            c.Font.Bold = True
        ElseIf TypeName(c) = "TextBox" Or TypeName(c) = "ComboBox" Then
            c.BackColor = RGB(255, 255, 255)
        End If
        On Error GoTo Failed
        n = c.Name
        If InStr(iconNames, "|" & n & "|") > 0 Then
            c.Left = rightX + countMoved * 73
            c.Top = 180
            c.Width = 66
            c.Height = 40
            countMoved = countMoved + 1
        End If
        If TypeName(c) = "Frame" Then
            If InStr(1, c.Caption, "Batch PowerClip", vbTextCompare) > 0 Then
                c.Left = rightX: c.Top = 12
                c.Width = 220
            End If
        End If
    Next c
    d.Width = 1050
    d.Height = 620
    d.Caption = "Card Maker | Production Studio"
    d.Font.Name = "Segoe UI"
    d.Font.Size = 9
    d.BackColor = RGB(246, 249, 252)
    For Each fr In d.Controls
        On Error Resume Next
        fr.Font.Name = "Segoe UI"
        fr.Font.Size = 9
        If TypeName(fr) = "Frame" Then
            fr.BackColor = RGB(255, 255, 255)
            fr.ForeColor = RGB(30, 47, 74)
        End If
        On Error GoTo Failed
    Next fr
    MsgBox "First-pass designer controls updated. Save and export Card_Maker.frm and Card_Maker.frx together." & vbCrLf & _
           "The runtime layout is handled by Card_Maker.frm. Check all control positions before using.", vbInformation
    Exit Sub
Failed:
    MsgBox "Designer update stopped: " & Err.Description & vbCrLf & _
           "Use a backup and ensure VBA project access is enabled.", vbExclamation
End Sub

Private Function FindFrameByCaption(ByVal formDesigner As Object, ByVal caption As String) As Object
    Dim c As Object
    For Each c In formDesigner.Controls
        If TypeName(c) = "Frame" Then
            If StrComp(c.Caption, caption, vbTextCompare) = 0 Then
                Set FindFrameByCaption = c
                Exit Function
            End If
        End If
    Next c
End Function
