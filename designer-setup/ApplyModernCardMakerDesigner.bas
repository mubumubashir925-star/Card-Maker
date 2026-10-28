Attribute VB_Name = "CardMakerDesignerSetup"
Option Explicit

' Card Maker V2 - native MSForms designer layout pass.
' Run on a BACKUP of the VBA project; save and export both FRM and FRX.
' Does not delete or rename existing controls or their event handlers.
Public Sub ApplyModernCardMakerDesigner()
    Dim comp As Object, d As Object, workspace As Object, c As Object
    Dim fitNames As Variant, fitTitles As Variant, i As Long
    Dim rightX As Single, rightW As Single, stageW As Single
    Dim fitTop As Single, toolY As Single, workTop As Single
    Dim maxRight As Single, maxBottom As Single
    On Error GoTo Failed
    Set comp = Application.VBE.ActiveVBProject.VBComponents("Card_Maker")
    Set d = comp.Designer
    Set workspace = FindFrameByCaption(d, "Card Maker")
    If workspace Is Nothing Then Err.Raise vbObjectError + 100, , "Original Card Maker frame not found"

    d.Caption = "CARD MAKER PRO  |  CorelDRAW Production Studio"
    d.BackColor = RGB(246, 249, 252)
    d.Font.Name = "Segoe UI": d.Font.Size = 9

    workspace.Caption = "Card Maker"  ' Runtime locates this frame by its exact caption.
    workspace.BackColor = RGB(255, 255, 255)
    workspace.ForeColor = RGB(31, 49, 77)
    workspace.Font.Name = "Segoe UI": workspace.Font.Size = 9
    workspace.Font.Bold = True
    workspace.BorderStyle = 1

    ' Preserve original field coordinates and existing child controls.
    ' Place the PowerClip and fitting controls in a dedicated right-hand zone.
    stageW = 0: maxBottom = 0
    For Each c In workspace.Controls
        If c.Name <> "FitImgtoPCframe" And c.Name <> "FitTextToFrame" And c.Name <> "Stretch_toFit" Then
            If TypeName(c) <> "Frame" Then
                If c.Left + c.Width > stageW Then stageW = c.Left + c.Width
            End If
            If c.Top + c.Height > maxBottom Then maxBottom = c.Top + c.Height
        End If
    Next c
    If stageW < 300 Then stageW = 300
    rightX = stageW + 20: rightW = 236
    toolY = 42

    For Each c In workspace.Controls
        On Error Resume Next
        c.Font.Name = "Segoe UI"
        c.Font.Size = 9
        If TypeName(c) = "CommandButton" Then
            c.Font.Bold = True
            c.BackColor = RGB(227, 239, 249)
            c.ForeColor = RGB(31, 49, 77)
            c.TakeFocusOnClick = False
        ElseIf TypeName(c) = "TextBox" Or TypeName(c) = "ComboBox" Or TypeName(c) = "ListBox" Then
            c.BackColor = RGB(255, 255, 255)
            c.ForeColor = RGB(31, 49, 77)
        ElseIf TypeName(c) = "Frame" Then
            c.BackColor = RGB(255, 255, 255)
            c.ForeColor = RGB(31, 49, 77)
            If InStr(1, c.Caption, "Batch PowerClip", vbTextCompare) > 0 Then
                c.Left = rightX: c.Top = toolY
                c.Width = rightW
                toolY = c.Top + c.Height + 22
            End If
        End If
        Err.Clear
        On Error GoTo Failed
    Next c

    fitNames = Array("FitImgtoPCframe", "FitTextToFrame", "Stretch_toFit")
    fitTitles = Array("FIT IMAGE", "FIT TEXT", "STRETCH")
    For i = 0 To 2
        Set c = FindControlByName(workspace, CStr(fitNames(i)))
        If Not c Is Nothing Then
            c.Left = rightX + i * 79
            c.Top = toolY + 20
            c.Width = 72: c.Height = 40
            AddOrUpdateLabel workspace, "v2FitLabel" & CStr(i + 1), CStr(fitTitles(i)), _
                             c.Left, toolY, 72, 18, RGB(31, 49, 77), RGB(255, 255, 255), 8
        End If
    Next i
    AddOrUpdateLabel workspace, "v2RightHeading", "IMAGE  /  POWERCLIP TOOLS", _
                     rightX, 16, rightW, 20, RGB(31, 49, 77), RGB(255, 255, 255), 10
    AddOrUpdateLabel workspace, "v2AdvancedHeading", "+  QR / BARCODE  |  Advanced settings at runtime", _
                     12, IIf(maxBottom > toolY + 74, maxBottom + 14, toolY + 84), _
                     stageW + rightW + 8, 25, RGB(31, 49, 77), RGB(236, 243, 249), 9

    workspace.Width = rightX + rightW + 15
    If workspace.Height < maxBottom + 60 Then workspace.Height = maxBottom + 60
    If workspace.Height < toolY + 125 Then workspace.Height = toolY + 125
    d.Width = workspace.Left + workspace.Width + 28
    d.Height = workspace.Top + workspace.Height + 70

    ' These headings are designer-native; runtime layout is built by Card_Maker.frm.
    MsgBox "Designer pass applied. Save the VBA project, then export Card_Maker.frm AND Card_Maker.frx." & vbCrLf & _
           "This preserves all original named controls and handlers. Verify layout in CorelDRAW X7.", vbInformation
    Exit Sub
Failed:
    MsgBox "Designer rebuild could not finish: " & Err.Description & vbCrLf & _
           "Keep the backup and confirm VBA project access is enabled.", vbExclamation
End Sub

Private Function FindFrameByCaption(ByVal d As Object, ByVal caption As String) As Object
    Dim c As Object
    For Each c In d.Controls
        If TypeName(c) = "Frame" Then
            If StrComp(c.Caption, caption, vbTextCompare) = 0 Then
                Set FindFrameByCaption = c
                Exit Function
            End If
        End If
    Next c
End Function

Private Function FindControlByName(ByVal parent As Object, ByVal controlName As String) As Object
    Dim c As Object
    For Each c In parent.Controls
        If StrComp(c.Name, controlName, vbTextCompare) = 0 Then
            Set FindControlByName = c
            Exit Function
        End If
    Next c
End Function

Private Sub AddOrUpdateLabel(ByVal parent As Object, ByVal controlName As String, _
                             ByVal caption As String, ByVal x As Single, ByVal y As Single, _
                             ByVal w As Single, ByVal h As Single, ByVal fg As Long, _
                             ByVal bg As Long, ByVal fontSize As Single)
    Dim c As Object
    Set c = FindControlByName(parent, controlName)
    If c Is Nothing Then Set c = parent.Controls.Add("Forms.Label.1", controlName, True)
    c.Caption = caption
    c.Left = x: c.Top = y: c.Width = w: c.Height = h
    c.BackColor = bg: c.ForeColor = fg
    c.Font.Name = "Segoe UI": c.Font.Size = fontSize
    c.Font.Bold = True
End Sub
