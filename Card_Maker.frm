VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Card_Maker 
   Caption         =   "Card Maker"
   ClientHeight    =   12090
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9600
   OleObjectBlob   =   "Card_Maker.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Card_Maker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'=====================================================================
' PART A  -  QR / Barcode generators: constants and variables
'            (must stay at the top of the form code, above the first Sub)
'=====================================================================
Private Const GC_FONT As String = "Arial"     ' font of the human-readable text
Private Const GC_QR_ECC As Long = 1           ' QR error correction: 0=L  1=M  2=Q  3=H
Private Const GC_QR_QUIET As Long = 4         ' empty border around QR (in modules)
Private Const GC_C128_QUIET As Long = 10      ' empty border left/right of Code 128 (in modules)
Private Const GC_WHITE_BG As Boolean = True   ' draw a white box behind the code
Private Const GC_FAST_PAGES As Boolean = False ' True = experimental faster page copying (test first!)

Private qrSize As Long
Private qrMod(0 To 96, 0 To 96) As Boolean
Private qrFn(0 To 96, 0 To 96) As Boolean
Private Const GC_NAVY As Long = &H442A1F      ' header / text colour
Private Const GC_TEAL As Long = &H8F8F0E      ' accent colour (buttons, active page)
Private Const GC_RAIL As Long = &HF0E9E4      ' navigation rail
Private Const GC_LIGHT As Long = &HD1B39F     ' light text on the header
Private Const GC_MUTED As Long = &H82685A     ' grey status text

Private gcSkip(1 To 2) As String
Private gcLastError As String
Private gcFields As Collection
Private gcTab(0 To 4) As Collection           ' pages: 0 Cards, 1 QR, 2 PowerClip, 3 Arrange, 4 Text / Group
Private gcUiReady As Boolean
Private gcMf As Object                        ' the main frame
Private gcFrameH(0 To 4) As Single            ' frame height on each page
Private gcGeo(1 To 2, 1 To 4) As Double       ' QR position: centre X, centre Y, width, height (inches)
Private gcGeoOk(1 To 2) As Boolean            ' True when Position was set
Private gcBusy As Boolean
Private gcT0 As Double
Private gcFill As Object
Private gcTrack As Object
Private gcFtText As Object
Private gcChip As Object
Private SelectedShapeIndices As Collection    ' used by Capture / Group (Text / Group page)
Private SourcePage As Page
Private WithEvents nv1 As MSForms.Label
Private WithEvents nv2 As MSForms.Label
Private WithEvents nv3 As MSForms.Label
Private WithEvents nv4 As MSForms.Label
Private WithEvents nv5 As MSForms.Label
Private WithEvents goBtn As MSForms.Label
Private WithEvents d1Add As MSForms.CommandButton
Private WithEvents d1Pos As MSForms.CommandButton
Private WithEvents d2Add As MSForms.CommandButton
Private WithEvents d2Pos As MSForms.CommandButton















Private Sub BarcodeMove_Click()
Dim iCtr As Long
 

    For iCtr = 0 To ListBox1.ListCount - 1
        If ListBox1.Selected(iCtr) = True Then
           BarList.text = ListBox1.List(iCtr)
           ListBox1.RemoveItem (iCtr)
        End If
        Next iCtr
        
        BarcodeMove.Enabled = False
End Sub

Private Sub cmdGroupObj_Click()
    On Error Resume Next
    Dim doc As Document
    Dim pg As Page
    Dim shapeNumbers(1 To 10) As Integer
    Dim shapesToGroup(1 To 10) As Shape
    Dim sr As ShapeRange
    Dim i As Integer
    Dim totalPages As Integer
    Dim currentPage As Integer
    Dim shapesGrouped As Boolean
    
    ' Initialize document and page count
    Set doc = ActiveDocument
    doc.Unit = cdrMillimeter
    totalPages = doc.Pages.Count
    currentPage = 0
    
    ' Configure and display the status label
    gcSetStatus "Processing Page 0 of " & totalPages
    DoEvents ' Ensure initial UI update
    
    ' Optimize performance and disable events
    Optimization = True
    EventsEnabled = False
    ActiveDocument.SaveSettings
    ActiveDocument.PreserveSelection = False
    
    ' Group commands for undo/redo efficiency
    doc.BeginCommandGroup "Group Objects"
    
    ' Get input from user for the shapes to group
    shapeNumbers(1) = Val(ShapeForGrp1.value) ' Convert text to number
    shapeNumbers(2) = Val(ShapeForGrp2.value)
    shapeNumbers(3) = Val(ShapeForGrp3.value)
    shapeNumbers(4) = Val(ShapeForGrp4.value)
    shapeNumbers(5) = Val(ShapeForGrp5.value)
    shapeNumbers(6) = Val(ShapeForGrp6.value)
    shapeNumbers(7) = Val(ShapeForGrp7.value)
    shapeNumbers(8) = Val(ShapeForGrp8.value)
    shapeNumbers(9) = Val(ShapeForGrp9.value)
    shapeNumbers(10) = Val(ShapeForGrp10.value)
    
    ' Validate at least one valid shape number
    Dim validInput As Boolean
    validInput = False
    For i = 1 To 10
        If shapeNumbers(i) > 0 Then
            validInput = True
            Exit For
        End If
    Next i
    
    If Not validInput Then
        gcIdle
        MsgBox "Please enter at least one valid shape number.", vbExclamation
        ActiveDocument.PreserveSelection = True
        ActiveDocument.RestoreSettings
        EventsEnabled = True
        Optimization = False
        doc.EndCommandGroup
        ActiveWindow.Refresh
        Application.Refresh
        Exit Sub
    End If
    
    ' Flag to track if any shapes were grouped
    shapesGrouped = False
    
    ' Iterate through each page in the document
    For Each pg In doc.Pages
        currentPage = currentPage + 1
        pg.Activate
        
        ' Update status with current page progress
        gcSetStatus "Processing Page " & currentPage & " of " & totalPages
        DoEvents ' Refresh UI
        
        ' Initialize the ShapeRange object
        Set sr = ActivePage.Shapes.Range
        
        ' Clear any existing shapes in the ShapeRange to avoid conflicts
        sr.RemoveAll
        
        ' Set the specified shapes for grouping
        For i = 1 To 10
            If shapeNumbers(i) > 0 Then
                On Error Resume Next
                Set shapesToGroup(i) = pg.ActiveLayer.Shapes(shapeNumbers(i))
                On Error GoTo 0
                If Not shapesToGroup(i) Is Nothing Then
                    sr.Add shapesToGroup(i)
                End If
            End If
        Next i
        
        ' Group the shapes if any valid shapes were added
        If sr.Count > 1 Then ' Ensure at least 2 shapes to group
            sr.group
            shapesGrouped = True
        End If
    Next pg
    
    ' Clean up and reset environment
    gcIdle
    ActiveDocument.PreserveSelection = True
    ActiveDocument.RestoreSettings
    EventsEnabled = True
    Optimization = False
    doc.EndCommandGroup
    
    ' Refresh the UI
    ActiveWindow.Refresh
    Application.Refresh
    
    ' Show completion message
    If shapesGrouped Then
        MsgBox "Shapes grouped successfully on all pages.", vbInformation, "Operation Complete"
    Else
        MsgBox "No valid shapes found to group or insufficient shapes to form a group.", vbExclamation
    End If
End Sub

Private Sub cmdMoveBackOrFrnt_Click()
    On Error Resume Next
    Dim doc As Document
    Dim totalPages As Integer
    Dim currentPage As Integer
    Dim shapeNumber As Integer
    Dim shapesMoved As Boolean
    
    ' Initialize document and page count
    Set doc = ActiveDocument
    doc.Unit = cdrCentimeter
    totalPages = doc.Pages.Count
    currentPage = 0
    
    ' Configure and display the status label
    gcSetStatus "Processing Page 0 of " & totalPages
    DoEvents ' Ensure initial UI update
    
    ' Optimize performance and disable events
    Optimization = True
    EventsEnabled = False
    ActiveDocument.SaveSettings
    ActiveDocument.PreserveSelection = False
    
    ' Group commands for undo/redo efficiency
    doc.BeginCommandGroup "Move Shape Order"
    
    ' Get shape number from user
    shapeNumber = Val(MoveShapeBack.value)
    
    ' Validate the input
    If shapeNumber <= 0 Then
        gcIdle
        MsgBox "Invalid shape number. Please enter a valid shape number in MoveShapeBack.", vbExclamation
        ActiveDocument.PreserveSelection = True
        ActiveDocument.RestoreSettings
        EventsEnabled = True
        Optimization = False
        doc.EndCommandGroup
        ActiveWindow.Refresh
        Application.Refresh
        Exit Sub
    End If
    
    ' Initialize shapesMoved flag
    shapesMoved = False
    
    ' Check which option is selected and call the respective function
    If MoveToFront.value = True Then
        shapesMoved = OptMoveShapeToFrnt(shapeNumber)
    ElseIf MoveToBack.value = True Then
        shapesMoved = OptMoveShapeBack(shapeNumber)
    Else
        ' If neither option is selected, display a warning message
        gcIdle
        MsgBox "Please select an option (Move to Front or Move to Back).", vbExclamation
        ActiveDocument.PreserveSelection = True
        ActiveDocument.RestoreSettings
        EventsEnabled = True
        Optimization = False
        doc.EndCommandGroup
        ActiveWindow.Refresh
        Application.Refresh
        Exit Sub
    End If
    
    ' Clean up and reset environment
    gcIdle
    ActiveDocument.PreserveSelection = True
    ActiveDocument.RestoreSettings
    EventsEnabled = True
    Optimization = False
    doc.EndCommandGroup
    
    ' Refresh the UI
    ActiveWindow.Refresh
    Application.Refresh
    
    ' Show completion message
    If shapesMoved Then
        MsgBox "Shape number " & shapeNumber & " moved successfully on all pages. Note: The shape's index may change after this operation.", vbInformation, "Operation Complete"
    Else
        MsgBox "No valid shape found with number " & shapeNumber & " on any page.", vbExclamation
    End If
End Sub

Private Function OptMoveShapeToFrnt(shapeNumber As Integer) As Boolean
    Dim pg As Page
    Dim shapeToMove As Shape
    Dim currentPage As Integer
    Dim totalPages As Integer
    Dim moved As Boolean
    ' Dim debugMsg As String ' For optional debug logging
    
    ' Initialize page count and flag
    totalPages = ActiveDocument.Pages.Count
    currentPage = 0
    moved = False
    
    ' Iterate through each page in the document
    For Each pg In ActiveDocument.Pages
        ' Update page counter and status
        currentPage = currentPage + 1
        pg.Activate
        gcSetStatus "Processing Page " & currentPage & " of " & totalPages
        DoEvents ' Refresh UI
        
        ' Set the specified shape for moving
        On Error Resume Next
        Set shapeToMove = pg.ActiveLayer.Shapes(shapeNumber)
        On Error GoTo 0
        
        ' Check if the shape exists on the page
        If Not shapeToMove Is Nothing Then
            shapeToMove.OrderToFront
            moved = True
            ' Optional debug logging: Uncomment to log new index
            ' debugMsg = "Page " & currentPage & ": Shape moved to front, new index = " & pg.ActiveLayer.Shapes.Count - pg.ActiveLayer.Shapes.IndexOf(shapeToMove) + 1
            ' Debug.Print debugMsg
        End If
    Next pg
    
    ' Return whether any shapes were moved
    OptMoveShapeToFrnt = moved
End Function

Private Function OptMoveShapeBack(shapeNumber As Integer) As Boolean
    Dim pg As Page
    Dim shapeToMove As Shape
    Dim currentPage As Integer
    Dim totalPages As Integer
    Dim moved As Boolean
    ' Dim debugMsg As String ' For optional debug logging
    
    ' Initialize page count and flag
    totalPages = ActiveDocument.Pages.Count
    currentPage = 0
    moved = False
    
    ' Iterate through each page in the document
    For Each pg In ActiveDocument.Pages
        ' Update page counter and status
        currentPage = currentPage + 1
        pg.Activate
        gcSetStatus "Processing Page " & currentPage & " of " & totalPages
        DoEvents ' Refresh UI
        
        ' Set the specified shape for moving
        On Error Resume Next
        Set shapeToMove = pg.ActiveLayer.Shapes(shapeNumber)
        On Error GoTo 0
        
        ' Check if the shape exists on the page
        If Not shapeToMove Is Nothing Then
            shapeToMove.OrderToBack
            moved = True
            ' Optional debug logging: Uncomment to log new index
            ' debugMsg = "Page " & currentPage & ": Shape moved to back, new index = " & pg.ActiveLayer.Shapes.Count - pg.ActiveLayer.Shapes.IndexOf(shapeToMove) + 1
            ' Debug.Print debugMsg
        End If
    Next pg
    
    ' Return whether any shapes were moved
    OptMoveShapeBack = moved
End Function

Private Sub cmdMoveObj_Click()
    On Error Resume Next
    Dim doc As Document
    Dim pg As Page
    Dim shapeNumber As Integer
    Dim xValue As Double
    Dim yValue As Double
    Dim shapeToMove As Shape
    Dim totalPages As Integer
    Dim currentPage As Integer
    Dim shapesMoved As Boolean
    
    ' Initialize document and page count
    Set doc = ActiveDocument
    doc.Unit = cdrMillimeter
    totalPages = doc.Pages.Count
    currentPage = 0
    
    ' Configure and display the status label
    gcSetStatus "Processing Page 0 of " & totalPages
    DoEvents ' Ensure initial UI update
    
    ' Optimize performance and disable events
    Optimization = True
    EventsEnabled = False
    ActiveDocument.SaveSettings
    ActiveDocument.PreserveSelection = False
    
    ' Group commands for undo/redo efficiency
    doc.BeginCommandGroup "Move Objects"
    
    ' Get input from user
    shapeNumber = Val(MoveShape.value)
    xValue = Val(MyX.value)
    yValue = Val(MyY.value)
    
    ' Validate the input
    If shapeNumber <= 0 Then
        gcIdle
        MsgBox "Invalid shape number. Please enter a valid shape number.", vbExclamation
        ActiveDocument.PreserveSelection = True
        ActiveDocument.RestoreSettings
        EventsEnabled = True
        Optimization = False
        doc.EndCommandGroup
        ActiveWindow.Refresh
        Application.Refresh
        Exit Sub
    End If
    
    ' Flag to track if any shapes were moved
    shapesMoved = False
    
    ' Iterate through each page in the document
    For Each pg In doc.Pages
        currentPage = currentPage + 1
        pg.Activate
        
        ' Update status with current page progress
        gcSetStatus "Processing Page " & currentPage & " of " & totalPages
        DoEvents ' Refresh UI
        
        ' Set the specified shape for moving
        On Error Resume Next
        Set shapeToMove = pg.ActiveLayer.Shapes(shapeNumber)
        On Error GoTo 0
        
        ' Check if the shape with the specified number exists on the page
        If Not shapeToMove Is Nothing Then
            ' Move the shape to the specified X and Y coordinates
            shapeToMove.centerX = xValue
            shapeToMove.centerY = yValue
            shapesMoved = True
        End If
    Next pg
    
    ' Clean up and reset environment
    gcIdle
    ActiveDocument.PreserveSelection = True
    ActiveDocument.RestoreSettings
    EventsEnabled = True
    Optimization = False
    doc.EndCommandGroup
    
    ' Refresh the UI
    ActiveWindow.Refresh
    Application.Refresh
    
    ' Show completion message
    If shapesMoved Then
        MsgBox "Shape number " & shapeNumber & " moved successfully on all pages.", vbInformation, "Operation Complete"
    Else
        MsgBox "No valid shape found with number " & shapeNumber & ".", vbExclamation
    End If
End Sub

Private Sub cmdShapeResize_Click()
    On Error Resume Next
    Dim doc As Document
    Dim pg As Page
    Dim shapeNumber As Integer
    Dim width As Double
    Dim length As Double
    Dim shapeToResize As Shape
    Dim totalPages As Integer
    Dim currentPage As Integer
    Dim shapesResized As Boolean
    
    ' Initialize document and page count
    Set doc = ActiveDocument
    doc.Unit = cdrMillimeter
    totalPages = doc.Pages.Count
    currentPage = 0
    
    ' Configure and display the status label
    gcSetStatus "Processing Page 0 of " & totalPages
    DoEvents ' Ensure initial UI update
    
    ' Optimize performance and disable events
    Optimization = True
    EventsEnabled = False
    ActiveDocument.SaveSettings
    ActiveDocument.PreserveSelection = False
    
    ' Group commands for undo/redo efficiency
    doc.BeginCommandGroup "Resize Shape"
    
    ' Extract values
    shapeNumber = Val(TargShape.value)
    width = Val(WidthR.value)
    length = Val(HeightR.value)
    
    ' Validate the input
    If shapeNumber <= 0 Then
        gcIdle
        MsgBox "Invalid shape number. Please enter a valid shape number in TargShape.", vbExclamation
        ActiveDocument.PreserveSelection = True
        ActiveDocument.RestoreSettings
        EventsEnabled = True
        Optimization = False
        doc.EndCommandGroup
        ActiveWindow.Refresh
        Application.Refresh
        Exit Sub
    End If
    
    If width <= 0 Or length <= 0 Then
        gcIdle
        MsgBox "Invalid width or length. Please enter positive values in WidthR and HeightR.", vbExclamation
        ActiveDocument.PreserveSelection = True
        ActiveDocument.RestoreSettings
        EventsEnabled = True
        Optimization = False
        doc.EndCommandGroup
        ActiveWindow.Refresh
        Application.Refresh
        Exit Sub
    End If
    
    ' Initialize shapesResized flag
    shapesResized = False
    
    ' Iterate through each page in the document
    For Each pg In doc.Pages
        ' Update page counter and status
        currentPage = currentPage + 1
        pg.Activate
        gcSetStatus "Processing Page " & currentPage & " of " & totalPages
        DoEvents ' Refresh UI
        
        ' Set the specified shape for resizing
        On Error Resume Next
        Set shapeToResize = pg.ActiveLayer.Shapes(shapeNumber)
        On Error GoTo 0
        
        ' Check if the shape exists on the page
        If Not shapeToResize Is Nothing Then
            ' Store original dimensions to verify resize
            Dim origWidth As Double
            Dim origHeight As Double
            origWidth = shapeToResize.SizeWidth
            origHeight = shapeToResize.SizeHeight
            
            ' Set the reference point for the transformation to the center of the shape
            doc.ReferencePoint = cdrCenter
            ' Resize the shape
            shapeToResize.SetSize width, length
            
            ' Verify resize was successful by checking new dimensions
            If Abs(shapeToResize.SizeWidth - width) < 0.001 And Abs(shapeToResize.SizeHeight - length) < 0.001 Then
                shapesResized = True
            End If
            
            ' Optional debug logging: Uncomment to log new index and dimensions
            ' Dim debugMsg As String
            ' debugMsg = "Page " & currentPage & ": Shape resized, index = " & pg.ActiveLayer.Shapes.Count - pg.ActiveLayer.Shapes.IndexOf(shapeToResize) + 1 & ", Width = " & shapeToResize.SizeWidth & ", Height = " & shapeToResize.SizeHeight
            ' Debug.Print debugMsg
        End If
    Next pg
    
    ' Clean up and reset environment
    gcIdle
    ActiveDocument.PreserveSelection = True
    ActiveDocument.RestoreSettings
    EventsEnabled = True
    Optimization = False
    doc.EndCommandGroup
    
    ' Refresh the UI
    ActiveWindow.Refresh
    Application.Refresh
    
    ' Show completion message
    If shapesResized Then
        MsgBox "Shape number " & shapeNumber & " resized successfully on all pages.", vbInformation, "Operation Complete"
    Else
        MsgBox "No valid shape found with number " & shapeNumber & " on any page.", vbExclamation
    End If
End Sub



Private Sub gcX_Change()

End Sub

Private Sub TextCaseApply_Click()
    Dim shapeNumber As Integer
    Dim caseOption As String

    ' Get input from user
    shapeNumber = Val(TextFitToFrame.text)
    
    ' Validate shape number
    If shapeNumber <= 0 Then
        MsgBox "Please enter a valid shape number.", vbExclamation
        Exit Sub
    End If

    ' Determine selected case option
    If OptionButtonUpperCase.value Then
        caseOption = "U"
    ElseIf OptionButtonLowerCase.value Then
        caseOption = "L"
    ElseIf OptionButtonTitleCase.value Then
        caseOption = "T"
    ElseIf OptionButtonToggleCase.value Then
        caseOption = "C"
    Else
        MsgBox "Please select a case option.", vbExclamation
        Exit Sub
    End If

    ' Call the ChangeTextCase macro
    ChangeTextCase shapeNumber, caseOption
End Sub

Sub ChangeTextCase(shapeNumber As Integer, caseOption As String)
    Dim doc As Document
    Dim pg As Page
    Dim shapeToChange As Shape
    Dim textContent As String
    Dim totalPages As Integer
    Dim currentPage As Integer
    
    ' Initialize document and page count
    Set doc = ActiveDocument
    totalPages = doc.Pages.Count
    currentPage = 0
    
    ' Configure and display the status label
    gcSetStatus "Processing Page 0 of " & totalPages
    DoEvents ' Ensure initial UI update
    
    ' Optimize performance and disable events
    Optimization = True
    EventsEnabled = False
    ActiveDocument.SaveSettings
    ActiveDocument.PreserveSelection = False
    
    ' Group commands for undo/redo efficiency
    doc.BeginCommandGroup "Text Case Adjustment"
    
    ' Iterate through pages
    For Each pg In doc.Pages
        currentPage = currentPage + 1
        pg.Activate
        
        ' Update status with current page progress
        gcSetStatus "Processing Page " & currentPage & " of " & totalPages
        DoEvents ' Refresh UI
        
        ' Check if the shape with the specified number exists on the page
        On Error Resume Next
        Set shapeToChange = pg.Shapes(shapeNumber)
        On Error GoTo 0

        ' If the shape exists and is a text frame, change the text case
        If Not shapeToChange Is Nothing Then
            If shapeToChange.Type = cdrTextShape Then
                textContent = shapeToChange.text.Story
                
                ' Verify text content exists before processing
                If Len(Trim(textContent)) > 0 Then
                    Select Case UCase(caseOption)
                        Case "U"
                            shapeToChange.text.Story = UCase(textContent) ' Upper case
                        Case "L"
                            shapeToChange.text.Story = LCase(textContent) ' Lower case
                        Case "T"
                            shapeToChange.text.Story = StrConv(textContent, vbProperCase) ' Title case
                        Case "C"
                            shapeToChange.text.Story = ToggleCase(textContent) ' Toggle case
                    End Select
                End If
            End If
        End If
    Next pg
    
    ' Finalize command group
    doc.EndCommandGroup
    
    ' Clean up and reset environment
    gcIdle
    ActiveDocument.PreserveSelection = True
    ActiveDocument.RestoreSettings
    EventsEnabled = True
    Optimization = False
    
    ' Refresh the UI
    ActiveWindow.Refresh
    Application.Refresh
    
    ' Show completion message
    MsgBox "Changed text case for shape number " & shapeNumber & " on all pages.", vbInformation, "Operation Complete"
End Sub

' Function to toggle the case of a string
Function ToggleCase(ByVal text As String) As String
    Dim i As Integer
    Dim result As String
    result = ""
    
    If Len(Trim(text)) = 0 Then
        ToggleCase = text
        Exit Function
    End If
    
    For i = 1 To Len(text)
        If Mid(text, i, 1) = UCase(Mid(text, i, 1)) Then
            result = result & LCase(Mid(text, i, 1))
        Else
            result = result & UCase(Mid(text, i, 1))
        End If
    Next i
    
    ToggleCase = result
End Function



Private Sub DeleteDatainEveryPage_Click()
    Dim doc As Document
    Dim pg As Page
    Dim layer As layer
    Dim selectedShape As Shape
    Dim shapeOrder As Long
    Dim totalPages As Integer
    Dim currentPage As Integer
    
    ' Disable events and save settings
    Optimization = True
    EventsEnabled = False
    ActiveDocument.SaveSettings
    ActiveDocument.PreserveSelection = False
    
    ' Initialize document and page count
    Set doc = ActiveDocument
    totalPages = doc.Pages.Count
    currentPage = 0
    
    ' Check if any shape is selected
    If activeSelection.Shapes.Count = 0 Then
        MsgBox "Oops! Nothing’s selected. Please select an object and try again.", vbExclamation, "No Selection"
        GoTo Cleanup
    End If
    
    ' Get the first selected shape and its layer
    Set selectedShape = activeSelection.Shapes(1)
    Set layer = selectedShape.layer
    
    ' Find the order of the selected shape within the layer
    For shapeOrder = 1 To layer.Shapes.Count
        If layer.Shapes(shapeOrder) Is selectedShape Then
            Exit For
        End If
    Next shapeOrder
    
    ' Configure and display the status label
    gcSetStatus "Processing Page 0 of " & totalPages & " with shape order " & shapeOrder & "..."
    DoEvents ' Ensure initial UI update
    
    ' Ask for user confirmation
    ans = MsgBox("Do you want to delete all shapes with the same shape number on all pages?", vbQuestion + vbYesNo, "Confirm Deletion")
    If ans = vbNo Then
        MsgBox "Deletion cancelled.", vbInformation, "Cancelled"
        gcIdle
        GoTo Cleanup
    End If
    
    ' Group commands for undo/redo efficiency
    doc.BeginCommandGroup "Delete Shapes Across Pages"
    
    On Error GoTo Cleanup ' Ensure cleanup occurs even if an error happens
    
    ' Loop through all pages
    For Each pg In doc.Pages
        currentPage = currentPage + 1
        pg.Activate
        
        ' Update status with current page progress
        gcSetStatus "Processing Page " & currentPage & " of " & totalPages & " with shape order " & shapeOrder & "..."
        DoEvents ' Refresh UI
        
        ' Check if the layer exists on this page and has enough shapes
        If Not pg.Layers(layer.Name) Is Nothing Then
            Set layer = pg.Layers(layer.Name)
            If shapeOrder <= layer.Shapes.Count Then
                ' Delete the shape at the same order on this page
                layer.Shapes(shapeOrder).Delete
            End If
        End If
    Next pg
    
Cleanup:
    ' Finalize command group and restore settings
    doc.EndCommandGroup
    gcIdle
    ActiveDocument.PreserveSelection = True
    ActiveDocument.RestoreSettings
    EventsEnabled = True
    Optimization = False
    ActiveWindow.Refresh
    Application.Refresh
End Sub
Sub DeleteShapesByNumberOnAllPages(shapeNumber As Long)
    ' Iterate through all pages
    For Each currentPage In ActiveDocument.Pages
        ' Check if the shape number is within the range for the current page
        If shapeNumber <= currentPage.Shapes.Count Then
            ' Delete the shape with the specified shape number on the current page
            currentPage.Shapes(shapeNumber).Delete
        End If
    Next currentPage
End Sub

Private Sub GetFitToFrame_Click()
Dim selectedShape As Shape
    Dim layer As layer
    Dim shapeOrder As Long

    ' Check if any shape is selected
    If activeSelection.Shapes.Count > 0 Then
        ' Get the first selected shape
        Set selectedShape = activeSelection.Shapes(1)

        ' Get the active layer
        Set layer = selectedShape.layer

        ' Find the order of the selected shape within the layer
        For shapeOrderFTF = 1 To layer.Shapes.Count
            If layer.Shapes(shapeOrderFTF) Is selectedShape Then
                ' Display the shape Number in a message box
                TextFitToFrame.value = shapeOrderFTF
                Exit Sub
            End If
        Next shapeOrderFTF
    Else
        MsgBox "No shape selected.", vbExclamation, "No Selection"
    End If
End Sub

Private Sub GetShapeForGroup10_Click()
Dim selectedShape As Shape
    Dim layer As layer
    Dim shapeOrder As Long

    ' Check if any shape is selected
    If activeSelection.Shapes.Count > 0 Then
        ' Get the first selected shape
        Set selectedShape = activeSelection.Shapes(1)

        ' Get the active layer
        Set layer = selectedShape.layer

        ' Find the order of the selected shape within the layer
        For shapeOrderG = 1 To layer.Shapes.Count
            If layer.Shapes(shapeOrderG) Is selectedShape Then
                ' Display the shape Number in a message box
                ShapeForGrp10.value = shapeOrderG
                Exit Sub
            End If
        Next shapeOrderG
    Else
        MsgBox "No shape selected.", vbExclamation, "No Selection"
    End If
End Sub

Private Sub GetShapeForGrp1_Click()
Dim selectedShape As Shape
    Dim layer As layer
    Dim shapeOrder As Long

    ' Check if any shape is selected
    If activeSelection.Shapes.Count > 0 Then
        ' Get the first selected shape
        Set selectedShape = activeSelection.Shapes(1)

        ' Get the active layer
        Set layer = selectedShape.layer

        ' Find the order of the selected shape within the layer
        For shapeOrderG = 1 To layer.Shapes.Count
            If layer.Shapes(shapeOrderG) Is selectedShape Then
                ' Display the shape Number in a message box
                ShapeForGrp1.value = shapeOrderG
                Exit Sub
            End If
        Next shapeOrderG
    Else
        MsgBox "No shape selected.", vbExclamation, "No Selection"
    End If
End Sub

Private Sub GetShapeForGrp2_Click()
Dim selectedShape As Shape
    Dim layer As layer
    Dim shapeOrder As Long

    ' Check if any shape is selected
    If activeSelection.Shapes.Count > 0 Then
        ' Get the first selected shape
        Set selectedShape = activeSelection.Shapes(1)

        ' Get the active layer
        Set layer = selectedShape.layer

        ' Find the order of the selected shape within the layer
        For shapeOrderG = 1 To layer.Shapes.Count
            If layer.Shapes(shapeOrderG) Is selectedShape Then
                ' Display the shape Number in a message box
                ShapeForGrp2.value = shapeOrderG
                Exit Sub
            End If
        Next shapeOrderG
    Else
        MsgBox "No shape selected.", vbExclamation, "No Selection"
    End If
End Sub

Private Sub GetShapeForGrp3_Click()
Dim selectedShape As Shape
    Dim layer As layer
    Dim shapeOrder As Long

    ' Check if any shape is selected
    If activeSelection.Shapes.Count > 0 Then
        ' Get the first selected shape
        Set selectedShape = activeSelection.Shapes(1)

        ' Get the active layer
        Set layer = selectedShape.layer

        ' Find the order of the selected shape within the layer
        For shapeOrderG = 1 To layer.Shapes.Count
            If layer.Shapes(shapeOrderG) Is selectedShape Then
                ' Display the shape Number in a message box
                ShapeForGrp3.value = shapeOrderG
                Exit Sub
            End If
        Next shapeOrderG
    Else
        MsgBox "No shape selected.", vbExclamation, "No Selection"
    End If
End Sub

Private Sub GetShapeForGrp4_Click()
Dim selectedShape As Shape
    Dim layer As layer
    Dim shapeOrder As Long

    ' Check if any shape is selected
    If activeSelection.Shapes.Count > 0 Then
        ' Get the first selected shape
        Set selectedShape = activeSelection.Shapes(1)

        ' Get the active layer
        Set layer = selectedShape.layer

        ' Find the order of the selected shape within the layer
        For shapeOrderG = 1 To layer.Shapes.Count
            If layer.Shapes(shapeOrderG) Is selectedShape Then
                ' Display the shape Number in a message box
                ShapeForGrp4.value = shapeOrderG
                Exit Sub
            End If
        Next shapeOrderG
    Else
        MsgBox "No shape selected.", vbExclamation, "No Selection"
    End If
End Sub

Private Sub GetShapeForGrp5_Click()
Dim selectedShape As Shape
    Dim layer As layer
    Dim shapeOrder As Long

    ' Check if any shape is selected
    If activeSelection.Shapes.Count > 0 Then
        ' Get the first selected shape
        Set selectedShape = activeSelection.Shapes(1)

        ' Get the active layer
        Set layer = selectedShape.layer

        ' Find the order of the selected shape within the layer
        For shapeOrderG = 1 To layer.Shapes.Count
            If layer.Shapes(shapeOrderG) Is selectedShape Then
                ' Display the shape Number in a message box
                ShapeForGrp5.value = shapeOrderG
                Exit Sub
            End If
        Next shapeOrderG
    Else
        MsgBox "No shape selected.", vbExclamation, "No Selection"
    End If
End Sub

Private Sub GetShapeForGrp6_Click()
Dim selectedShape As Shape
    Dim layer As layer
    Dim shapeOrder As Long

    ' Check if any shape is selected
    If activeSelection.Shapes.Count > 0 Then
        ' Get the first selected shape
        Set selectedShape = activeSelection.Shapes(1)

        ' Get the active layer
        Set layer = selectedShape.layer

        ' Find the order of the selected shape within the layer
        For shapeOrderG = 1 To layer.Shapes.Count
            If layer.Shapes(shapeOrderG) Is selectedShape Then
                ' Display the shape Number in a message box
                ShapeForGrp6.value = shapeOrderG
                Exit Sub
            End If
        Next shapeOrderG
    Else
        MsgBox "No shape selected.", vbExclamation, "No Selection"
    End If
End Sub

Private Sub GetShapeForGrp7_Click()
Dim selectedShape As Shape
    Dim layer As layer
    Dim shapeOrder As Long

    ' Check if any shape is selected
    If activeSelection.Shapes.Count > 0 Then
        ' Get the first selected shape
        Set selectedShape = activeSelection.Shapes(1)

        ' Get the active layer
        Set layer = selectedShape.layer

        ' Find the order of the selected shape within the layer
        For shapeOrderG = 1 To layer.Shapes.Count
            If layer.Shapes(shapeOrderG) Is selectedShape Then
                ' Display the shape Number in a message box
                ShapeForGrp7.value = shapeOrderG
                Exit Sub
            End If
        Next shapeOrderG
    Else
        MsgBox "No shape selected.", vbExclamation, "No Selection"
    End If
End Sub

Private Sub GetShapeForGrp8_Click()
Dim selectedShape As Shape
    Dim layer As layer
    Dim shapeOrder As Long

    ' Check if any shape is selected
    If activeSelection.Shapes.Count > 0 Then
        ' Get the first selected shape
        Set selectedShape = activeSelection.Shapes(1)

        ' Get the active layer
        Set layer = selectedShape.layer

        ' Find the order of the selected shape within the layer
        For shapeOrderG = 1 To layer.Shapes.Count
            If layer.Shapes(shapeOrderG) Is selectedShape Then
                ' Display the shape Number in a message box
                ShapeForGrp8.value = shapeOrderG
                Exit Sub
            End If
        Next shapeOrderG
    Else
        MsgBox "No shape selected.", vbExclamation, "No Selection"
    End If
End Sub

Private Sub GetShapeForGrp9_Click()
Dim selectedShape As Shape
    Dim layer As layer
    Dim shapeOrder As Long

    ' Check if any shape is selected
    If activeSelection.Shapes.Count > 0 Then
        ' Get the first selected shape
        Set selectedShape = activeSelection.Shapes(1)

        ' Get the active layer
        Set layer = selectedShape.layer

        ' Find the order of the selected shape within the layer
        For shapeOrderG = 1 To layer.Shapes.Count
            If layer.Shapes(shapeOrderG) Is selectedShape Then
                ' Display the shape Number in a message box
                ShapeForGrp9.value = shapeOrderG
                Exit Sub
            End If
        Next shapeOrderG
    Else
        MsgBox "No shape selected.", vbExclamation, "No Selection"
    End If
End Sub

Private Sub GetShapeNoMoveObj_Click()
Dim selectedShape As Shape
    Dim layer As layer
    Dim shapeOrder As Long

    ' Check if any shape is selected
    If activeSelection.Shapes.Count > 0 Then
        ' Get the first selected shape
        Set selectedShape = activeSelection.Shapes(1)

        ' Get the active layer
        Set layer = selectedShape.layer

        ' Find the order of the selected shape within the layer
        For shapeOrderR = 1 To layer.Shapes.Count
            If layer.Shapes(shapeOrderR) Is selectedShape Then
                ' Display the shape Number in a message box
                MoveShape.value = shapeOrderR
                Exit Sub
            End If
        Next shapeOrderR
    Else
        MsgBox "No shape selected.", vbExclamation, "No Selection"
    End If
End Sub

Private Sub ResetGrouping_Click()

    Me.ShapeForGrp1.value = ""
    Me.ShapeForGrp2.value = ""
    Me.ShapeForGrp3.value = ""
    Me.ShapeForGrp4.value = ""
    Me.ShapeForGrp5.value = ""
    Me.ShapeForGrp6.value = ""
    Me.ShapeForGrp7.value = ""
    Me.ShapeForGrp8.value = ""
    Me.ShapeForGrp9.value = ""
    Me.ShapeForGrp10.value = ""

End Sub

Private Sub Stretch_toFit_Click()
    Dim s As Shape
    Dim srPowerClipped As ShapeRange
    Dim targetWidth As Double, targetHeight As Double
    Dim doc As Document
    Dim pg As Page
    Dim layer As layer
    Dim shapeIndex As Long
    Dim totalPages As Integer
    Dim currentPage As Integer
    
    ' Disable events and save settings
    Optimization = True
    EventsEnabled = False
    ActiveDocument.SaveSettings
    ActiveDocument.PreserveSelection = False
    
    ' Initialize document and page count
    Set doc = ActiveDocument
    totalPages = doc.Pages.Count
    currentPage = 0
    
    ' Check if any shape is selected
    If activeSelection.Shapes.Count = 0 Then
        MsgBox "Oops! Nothing’s selected. Please select a powerclip frame and try again.", vbExclamation, "No Selection"
        GoTo Cleanup
    End If
    
    ' Get the selected shape and its layer
    Set s = activeSelection.Shapes(1)
    Set layer = s.layer
    
    ' Find the index of the selected shape within its layer
    For shapeIndex = 1 To layer.Shapes.Count
        If layer.Shapes(shapeIndex) Is s Then
            Exit For
        End If
    Next shapeIndex
    
    ' Configure and display the status label
    gcSetStatus "Processing Page 0 of " & totalPages & " with index " & shapeIndex & "..."
    DoEvents ' Ensure initial UI update
    
    ' User confirmation
    ans = MsgBox("Going to stretch content to fit frames with index " & shapeIndex & " on all pages," & vbNewLine & _
        "Do you want to continue?", vbQuestion + vbOKCancel, "Please confirm...")
    If ans = vbCancel Then
        gcIdle
        GoTo Cleanup
    End If
    
    ' Start a command group for single undo
    doc.BeginCommandGroup "Stretch to Fit Across Pages"
    
    On Error GoTo Cleanup ' Ensure cleanup occurs even if an error happens
    
    ' Loop through all pages
    For Each pg In doc.Pages
        currentPage = currentPage + 1
        pg.Activate
        
        ' Update status with current page progress
        gcSetStatus "Processing Page " & currentPage & " of " & totalPages & " with index " & shapeIndex & "..."
        DoEvents ' Refresh UI
        
        ' Check if the layer exists on this page and has enough shapes
        If Not pg.Layers(layer.Name) Is Nothing Then
            Set layer = pg.Layers(layer.Name)
            If shapeIndex <= layer.Shapes.Count Then
                ' Get the shape at the same index on this page
                Set s = layer.Shapes(shapeIndex)
                
                ' Get the PowerClip content of the shape
                Set srPowerClipped = s.PowerClip.Shapes.FindShapes()
                
                ' Set target dimensions to match the shape
                targetWidth = s.SizeWidth
                targetHeight = s.SizeHeight
                
                ' Stretch the PowerClip content to fit the shape, if content exists
                If srPowerClipped.Count > 0 Then
                    srPowerClipped.SetSize targetWidth, targetHeight
                    srPowerClipped.SetPositionEx cdrCenter, s.centerX, s.centerY
                End If
            End If
        End If
    Next pg
    
Cleanup:
    ' Finalize command group and restore settings
    doc.EndCommandGroup
    gcIdle
    ActiveDocument.PreserveSelection = True
    ActiveDocument.RestoreSettings
    EventsEnabled = True
    Optimization = False
    ActiveWindow.Refresh
    Application.Refresh
End Sub

Private Sub UngrpAll_Click()
   
   Optimization = True
EventsEnabled = False
ActiveDocument.SaveSettings
ActiveDocument.PreserveSelection = False

   ans = MsgBox("Going to Ungroup everything in each Page of this Document," & vbNewLine & _
   "Do you want to Continue?", vbQuestion + vbOKCancel, "Please confirm...")
    
        If ans = vbCancel Then
           ActiveDocument.RestoreSettings
            EventsEnabled = True
            Optimization = False
            ActiveDocument.EndCommandGroup
            ActiveWindow.Refresh
            Application.Refresh
            Exit Sub
        End If
    
    ' Iterate through each page in the document
    For Each pg In ActiveDocument.Pages
        ' Activate the current page
        pg.Activate

        ' Ungroup all objects on the page
        ActivePage.Shapes.All.Ungroup

Next
MsgBox "Ungrouped Successfully", vbInformation, "Ungrouping"

ActiveDocument.PreserveSelection = True
ActiveDocument.RestoreSettings
EventsEnabled = True
Optimization = False
ActiveDocument.EndCommandGroup
ActiveWindow.Refresh
Application.Refresh
End Sub

Private Sub createCards_Click()
    ' Disable unnecessary events and optimize performance
    Optimization = True
    EventsEnabled = False
    ActiveDocument.SaveSettings
    ActiveDocument.PreserveSelection = False
    ActiveDocument.Unit = cdrInch   ' saved positions/sizes are in inches (RestoreSettings returns your unit)

    Dim filePath As String
    Dim s As Shape
    Dim import As StructImportOptions
    Dim objFSO As Object
    Dim objStream As Object ' Changed to Object for Unicode handling
    Dim totalLines As Long
    Dim LineNum As Long
    Dim dummyCount As Long ' Counter for dummy images
    Dim dummyPages As String ' Comma-separated page numbers
    
    SourceType = ".jpg"     ' Include the dot, use lower case
    cdrType = cdrJPG
    dummyCount = 0 ' Initialize dummy image counter
    dummyPages = "" ' Initialize page number list
    gcSkip(1) = ""
    gcSkip(2) = ""
    gcLastError = ""

    ' Configure and display the status label
    gcSetStatus "Processing Page 0 of 0"
    DoEvents

    ' Validation checks
    tpg = ActiveDocument.Pages.Count
    If tpg > 1 Then
        MsgBox "There are 2 or more pages already exists. Make sure only one design page exists.", vbOKOnly, "Pages Alert - Only 1 Design Page Required"
        GoTo Cleanup
    End If
        
    If ListBox1.ListCount = 0 Then
        MsgBox "No TEXT File selected. Why kidding me (:.", vbOKOnly, "Database Alert - Select CSV file first"
        GoTo Cleanup
    End If
        
    If ipX.value = 0 And imgCheck.value = True Then
        MsgBox "You have not set the position for Image to place, Plz add them too! (: ", vbOKOnly
        GoTo Cleanup
    End If
        
    filePath = folderpath.value
    Dim cmT0 As Double
    cmT0 = Timer
    
    Dim gcS As Long, gcVal As String, gcDone As Boolean
    For gcS = 1 To 2
        If gcIsOn(gcS) Then
            If Trim$(gcCtl(gcS, "Col").Value) = "" Then
                MsgBox "QR/Barcode " & gcS & " is ticked but its content is empty. Choose a field and click Add.", vbCritical
                GoTo Cleanup
            End If
            If Not gcGeoOk(gcS) Then
                MsgBox "QR/Barcode " & gcS & " is ticked but its Position is not set. Select the box shape and click Position.", vbCritical
                GoTo Cleanup
            End If
        End If
    Next gcS
    Dim gcOn(1 To 2) As Boolean, gcTpl(1 To 2) As String, gcTyp(1 To 2) As String, gcTxt(1 To 2) As Boolean
    For gcS = 1 To 2
        gcOn(gcS) = gcIsOn(gcS)
        If gcOn(gcS) Then
            gcTpl(gcS) = gcCtl(gcS, "Col").Value
            gcTyp(gcS) = gcCtl(gcS, "Type").Value
            gcTxt(gcS) = CBool(gcCtl(gcS, "Text").Value)
        End If
    Next gcS
    'Check duplicate placeholders in design
    If Not CheckDuplicatePlaceholders() Then
        MsgBox "Process cancelled due to duplicate placeholders.", vbCritical
        GoTo Cleanup
    End If
    
    ' Get total lines in file for status
    Set objStream = OpenTextFileAuto(filePath)
    totalLines = 0
    Do While Not objStream.AtEndOfStream
        objStream.SkipLine
        totalLines = totalLines + 1
    Loop
    objStream.Close
    
    ' Reopen file for processing
    Set objStream = OpenTextFileAuto(filePath)
    LineNum = 0
            
    Do Until objStream.AtEndOfStream
        DupPage
        ActiveDocument.Pages(ActivePage.index - 1).Activate
        LineNum = LineNum + 1
    
        gcSetStatus "Processing Page " & (LineNum - 1) & " of " & (totalLines - 1)
        DoEvents
        
        Dim csdt1 As String
        Dim csdt2 As String
    
        LineFromFile = objStream.ReadLine
        
        ' Strip BOM character if present (Only on the first line)
        If LineNum = 1 Then
            If Len(LineFromFile) > 0 Then
                If AscW(Left(LineFromFile, 1)) = -257 Then LineFromFile = Mid(LineFromFile, 2)
            End If
        End If
        
        If LineNum > 1 Then
            Dim LineItems() As String
            If Trim(LineFromFile) = "" Then
                Debug.Print "Blank line at LineNum: " & LineNum & ", using dummy images and empty text"
                LineItems = Array("") ' Array with one empty string for text replacement
            Else
                LineItems = Split(LineFromFile, vbTab)
                If UBound(LineItems) < 0 Then
                    Debug.Print "Invalid line (no data after split) at LineNum: " & LineNum & ", using dummy images and empty text"
                    LineItems = Array("") ' Array with one empty string for text replacement
                End If
            End If
    
            ' Define dummy image path once
            Dim dummyImgPath As String
            dummyImgPath = "C:\Program Files\Corel\Card Maker\100.jpg"
            
            ' Adding first image
            If imgCheck.value = True Then
                If ImgList.value = "" Then
                    MsgBox "You have not added Image data field ", vbCritical
                    objStream.Close
                    GoTo Cleanup
                Else
                    imgPath = ImgList.value
                    numI = Val(Left(imgPath, 2))
                    Dim imgstr As String
                    Dim useDummy As Boolean: useDummy = False
                    If numI <= UBound(LineItems) And numI >= LBound(LineItems) Then
                        imgstr = LineItems(numI)
                        If Trim(imgstr) = "" Or Not FileExists(imgstr) Then
                            Debug.Print "Image not found or empty for index " & numI & " at LineNum: " & LineNum & ", using dummy"
                            useDummy = True
                        End If
                    Else
                        Debug.Print "Invalid index " & numI & " for first image at LineNum: " & LineNum & ", using dummy"
                        useDummy = True
                    End If

                    Dim impflt As Object
                    If Not useDummy Then
                        Set impflt = ActiveLayer.ImportEx(imgstr, cdrType)
                        impflt.Finish
                    Else
                        If FileExists(dummyImgPath) Then
                            Set impflt = ActiveLayer.ImportEx(dummyImgPath, cdrType)
                            impflt.Finish
                            Debug.Print "Dummy image placed for first image at LineNum: " & LineNum
                            dummyCount = dummyCount + 1
                            If dummyPages = "" Then
                                dummyPages = CStr(LineNum - 1)
                            ElseIf InStr(dummyPages, CStr(LineNum - 1)) = 0 Then
                                dummyPages = dummyPages & ", " & (LineNum - 1)
                            End If
                        Else
                            MsgBox "Dummy image file not found at " & dummyImgPath & ".", vbCritical
                            ActiveDocument.Pages(1).Delete
                            objStream.Close
                            GoTo Cleanup
                        End If
                    End If
                            
                    Dim s1 As Shape
                    Set s1 = ActiveShape
        
                    If iWch.value = True Then
                        s1.SetSize iwth.value
                    Else
                        s1.SetSize , ihth.value
                    End If
                    s1.centerX = ipX
                    s1.centerY = ipY
                End If
            End If
                    
            ' Adding 2nd image
            If imgCheck1.value = True Then
                If ImgList1.value = "" Then
                    MsgBox "You have not added Image data field ", vbCritical
                    objStream.Close
                    GoTo Cleanup
                Else
                    img2path = ImgList1.value
                    numC = Val(Left(img2path, 2))
                    Dim imcstr As String
                    Dim useDummy2 As Boolean: useDummy2 = False
                    If numC <= UBound(LineItems) And numC >= LBound(LineItems) Then
                        imcstr = LineItems(numC)
                        If Trim(imcstr) = "" Or Not FileExists(imcstr) Then
                            Debug.Print "Image not found or empty for index " & numC & " at LineNum: " & LineNum & ", using dummy"
                            useDummy2 = True
                        End If
                    Else
                        Debug.Print "Invalid index " & numC & " for second image at LineNum: " & LineNum & ", using dummy"
                        useDummy2 = True
                    End If
                        
                    Dim impflt2 As Object
                    If Not useDummy2 Then
                        Set impflt2 = ActiveLayer.ImportEx(imcstr, cdrType)
                        impflt2.Finish
                    Else
                        If FileExists(dummyImgPath) Then
                            Set impflt2 = ActiveLayer.ImportEx(dummyImgPath, cdrType)
                            impflt2.Finish
                            Debug.Print "Dummy image placed for second image at LineNum: " & LineNum
                            dummyCount = dummyCount + 1
                            If dummyPages = "" Then
                                dummyPages = CStr(LineNum - 1)
                            ElseIf InStr(dummyPages, CStr(LineNum - 1)) = 0 Then
                                dummyPages = dummyPages & ", " & (LineNum - 1)
                            End If
                        Else
                            MsgBox "Dummy image file not found at " & dummyImgPath & ".", vbCritical
                            ActiveDocument.Pages(1).Delete
                            objStream.Close
                            GoTo Cleanup
                        End If
                    End If

                    Dim s3 As Shape
                    Set s3 = ActiveShape

                    If iWch1.value = True Then
                        s3.SetSize iwth1.value
                    Else
                        s3.SetSize , ihth1.value
                    End If
                    s3.centerX = ipX1
                    s3.centerY = ipY1
                End If
            End If
                        
            ' Adding Barcode
            If bcCheck.value Then
                If BarList.value = "" Then
                    MsgBox "You have not added Image data field ", vbCritical
                    objStream.Close
                    GoTo Cleanup
                Else
                    bcpath = BarList.value
                    numB = Val(Left(bcpath, 2))
                    Dim barcodestr As String
                    Dim useDummy3 As Boolean: useDummy3 = False
                    If numB <= UBound(LineItems) And numB >= LBound(LineItems) Then
                        barcodestr = LineItems(numB)
                        If Trim(barcodestr) = "" Or Not FileExists(barcodestr) Then
                            Debug.Print "Barcode not found or empty for index " & numB & " at LineNum: " & LineNum & ", using dummy"
                            useDummy3 = True
                        End If
                    Else
                        Debug.Print "Invalid index " & numB & " for barcode at LineNum: " & LineNum & ", using dummy"
                        useDummy3 = True
                    End If

                    Dim impflt3 As Object
                    If Not useDummy3 Then
                        Set impflt3 = ActiveLayer.ImportEx(barcodestr, cdrType)
                        impflt3.Finish
                    Else
                        If FileExists(dummyImgPath) Then
                            Set impflt3 = ActiveLayer.ImportEx(dummyImgPath, cdrType)
                            impflt3.Finish
                            Debug.Print "Dummy image placed for barcode at LineNum: " & LineNum
                            dummyCount = dummyCount + 1
                            If dummyPages = "" Then
                                dummyPages = CStr(LineNum - 1)
                            ElseIf InStr(dummyPages, CStr(LineNum - 1)) = 0 Then
                                dummyPages = dummyPages & ", " & (LineNum - 1)
                            End If
                        Else
                            MsgBox "Dummy image file not found at " & dummyImgPath & ".", vbCritical
                            ActiveDocument.Pages(1).Delete
                            objStream.Close
                            GoTo Cleanup
                        End If
                    End If
                    
                    Dim b1 As Shape
                    Set b1 = ActiveShape
                                       
                    If bWch.value = True Then
                        b1.SetSize bwth.value
                    Else
                        b1.SetSize , bhth.value
                        bHch.value = True
                    End If
                    b1.centerX = bcX
                    b1.centerY = bcY
                        
                    ' Adding 3rd image
                    If imgCheck2.value = True Then
                        If ImgList2.value = "" Then
                            MsgBox "You have not added Image data field ", vbCritical
                            objStream.Close
                            GoTo Cleanup
                        Else
                            img3path = ImgList2.value
                            numD = Val(Left(img3path, 2))
                            Dim imdstr As String
                            Dim useDummy4 As Boolean: useDummy4 = False
                            If numD <= UBound(LineItems) And numD >= LBound(LineItems) Then
                                imdstr = LineItems(numD)
                                If Trim(imdstr) = "" Or Not FileExists(imdstr) Then
                                    Debug.Print "Image not found or empty for index " & numD & " at LineNum: " & LineNum & ", using dummy"
                                    useDummy4 = True
                                End If
                            Else
                                Debug.Print "Invalid index " & numD & " for third image at LineNum: " & LineNum & ", using dummy"
                                useDummy4 = True
                            End If

                            Dim impflt4 As Object
                            If Not useDummy4 Then
                                Set impflt4 = ActiveLayer.ImportEx(imdstr, cdrType)
                                impflt4.Finish
                            Else
                                If FileExists(dummyImgPath) Then
                                    Set impflt4 = ActiveLayer.ImportEx(dummyImgPath, cdrType)
                                    impflt4.Finish
                                    Debug.Print "Dummy image placed for third image at LineNum: " & LineNum
                                    dummyCount = dummyCount + 1
                                    If dummyPages = "" Then
                                        dummyPages = CStr(LineNum - 1)
                                    ElseIf InStr(dummyPages, CStr(LineNum - 1)) = 0 Then
                                        dummyPages = dummyPages & ", " & (LineNum - 1)
                                    End If
                                Else
                                    MsgBox "Dummy image file not found at " & dummyImgPath & ".", vbCritical
                                    ActiveDocument.Pages(1).Delete
                                    objStream.Close
                                    GoTo Cleanup
                                End If
                            End If
                    
                            Dim s4 As Shape
                            Set s4 = ActiveShape
                    
                            If iWch1.value = True Then
                                s4.SetSize iwth2.value
                            Else
                                s4.SetSize , ihth2.value
                            End If
                            s4.centerX = ipX2
                            s4.centerY = ipY2
                        End If
                    End If
                End If
            End If
            
            ' Adding generated QR / Barcodes (generator 1 and 2)
            For gcS = 1 To 2
                If gcOn(gcS) Then
                    gcVal = gcComposeValue(gcTpl(gcS), LineItems)
                    gcDone = False
                    If Trim$(gcVal) <> "" Then
                        gcDone = gcAddCode(gcVal, gcTyp(gcS), gcGeo(gcS, 1), gcGeo(gcS, 2), gcGeo(gcS, 3), gcGeo(gcS, 4), _
                                           gcTxt(gcS), "GeneratedCode" & gcS)
                    End If
                    If Not gcDone Then gcSkip(gcS) = gcSkip(gcS) & (LineNum - 1) & ", "
                End If
            Next gcS
                    
            ' Import text
            Dim txtcount As Integer
            Dim demotxt As String, origtxt As String
            Dim index As Integer

            txtcount = UBound(LineItems)
            ActiveDocument.BeginCommandGroup "Text Translate"

            For Each s In ActiveDocument.ActivePage.Shapes
                If s.Type = cdrTextShape Then
                    demotxt = s.text.Story
                    origtxt = demotxt
                    Do While InStr(demotxt, "<") > 0 And InStr(demotxt, ">") > 0
                        Dim startIndex As Integer
                        Dim endIndex As Integer
                        Dim replaceText As String
                        startIndex = InStr(1, demotxt, "<") + 1
                        endIndex = InStr(startIndex, demotxt, ">")
                        If endIndex = 0 Then endIndex = Len(demotxt)
                        index = Val(Mid(demotxt, startIndex, endIndex - startIndex))
                        replaceText = Mid(demotxt, startIndex - 1, endIndex - startIndex + 2)
                        If index >= 0 And index <= txtcount Then
                            demotxt = Replace(demotxt, replaceText, LineItems(index))
                        Else
                            demotxt = Replace(demotxt, replaceText, "")
                            Debug.Print "Replaced invalid text index " & index & " with empty string at LineNum: " & LineNum
                        End If
                    Loop
                    If demotxt <> origtxt Then s.text.Story = demotxt
                End If
            Next s
            ActiveDocument.EndCommandGroup
        End If
    
NextLoop:
        ActiveDocument.Pages(ActivePage.index + 1).Activate
    Loop
    
    objStream.Close
    ActiveDocument.Pages(1).Delete
    ActiveDocument.Pages(ActiveDocument.Pages.Count).Delete
    
    ' Update DummyList label with dummy image usage
    If dummyCount > 0 Then
        DummyList.Caption = "Dummy images on pages: " & dummyPages
    Else
        DummyList.Caption = "No dummy images placed."
    End If
    DummyList.Caption = DummyList.Caption & "  Done in " & Format(Timer - cmT0, "0") & " s."
    DoEvents
    If gcSkip(1) <> "" Or gcSkip(2) <> "" Then
        Dim gcMsg As String
        gcMsg = ""
        For gcS = 1 To 2
            If gcSkip(gcS) <> "" Then gcMsg = gcMsg & "QR/Barcode " & gcS & " was NOT created on page(s): " & _
                                              Left$(gcSkip(gcS), Len(gcSkip(gcS)) - 2) & vbCrLf
        Next gcS
        MsgBox gcMsg & "(empty value or a value that cannot be encoded)" & gcLastError, vbExclamation
    End If
    
Cleanup:
    gcIdle
    ActiveDocument.PreserveSelection = True
    ActiveDocument.RestoreSettings
    EventsEnabled = True
    Optimization = False
    ActiveDocument.EndCommandGroup
    ActiveWindow.Refresh
    Application.Refresh
End Sub

'=====================================================================
' PART B  -  QR / Barcode encoders, drawing, tabbed layout, generator panels
'=====================================================================
'---------------- QR code + Code 128 encoders (tested) ----------------

Private Function gcShl(ByVal v As Long, ByVal n As Long) As Long
    gcShl = v * CLng(2 ^ n)
End Function

Private Function gcShr(ByVal v As Long, ByVal n As Long) As Long
    gcShr = v \ CLng(2 ^ n)
End Function

Private Function gcBit(ByVal v As Long, ByVal i As Long) As Long
    gcBit = gcShr(v, i) And 1
End Function

'---------- text -> UTF-8 bytes ----------
Private Function gcUtf8(ByVal s As String, ByRef b() As Byte) As Long
    Dim i As Long, c As Long, c2 As Long, n As Long
    ReDim b(0 To Len(s) * 4 + 3)
    n = 0
    i = 1
    Do While i <= Len(s)
        c = AscW(Mid$(s, i, 1))
        If c < 0 Then c = c + 65536
        i = i + 1
        If c >= &HD800& And c <= &HDBFF& And i <= Len(s) Then
            c2 = AscW(Mid$(s, i, 1))
            If c2 < 0 Then c2 = c2 + 65536
            If c2 >= &HDC00& And c2 <= &HDFFF& Then
                c = &H10000 + ((c - &HD800&) * 1024) + (c2 - &HDC00&)
                i = i + 1
            End If
        End If
        If c < 128 Then
            b(n) = c: n = n + 1
        ElseIf c < 2048 Then
            b(n) = &HC0& Or (c \ 64): n = n + 1
            b(n) = &H80& Or (c And &H3F&): n = n + 1
        ElseIf c < 65536 Then
            b(n) = &HE0& Or (c \ 4096): n = n + 1
            b(n) = &H80& Or ((c \ 64) And &H3F&): n = n + 1
            b(n) = &H80& Or (c And &H3F&): n = n + 1
        Else
            b(n) = &HF0& Or (c \ 262144): n = n + 1
            b(n) = &H80& Or ((c \ 4096) And &H3F&): n = n + 1
            b(n) = &H80& Or ((c \ 64) And &H3F&): n = n + 1
            b(n) = &H80& Or (c And &H3F&): n = n + 1
        End If
    Loop
    gcUtf8 = n
End Function

'---------- QR tables (versions 1-10) ----------
Private Function gcEccPerBlock(ByVal ecl As Long, ByVal ver As Long) As Long
    Dim t As Variant
    Select Case ecl
        Case 0: t = Array(0, 7, 10, 15, 20, 26, 18, 20, 24, 30, 18, 20, 24, 26, 30, 22, 24, 28, 30, 28, 28)
        Case 1: t = Array(0, 10, 16, 26, 18, 24, 16, 18, 22, 22, 26, 30, 22, 22, 24, 24, 28, 28, 26, 26, 26)
        Case 2: t = Array(0, 13, 22, 18, 26, 18, 24, 18, 22, 20, 24, 28, 26, 24, 20, 30, 24, 28, 28, 26, 30)
        Case Else: t = Array(0, 17, 28, 22, 16, 22, 28, 26, 26, 24, 28, 24, 28, 22, 24, 24, 30, 28, 28, 26, 28)
    End Select
    gcEccPerBlock = t(LBound(t) + ver)
End Function

Private Function gcNumBlocks(ByVal ecl As Long, ByVal ver As Long) As Long
    Dim t As Variant
    Select Case ecl
        Case 0: t = Array(0, 1, 1, 1, 1, 1, 2, 2, 2, 2, 4, 4, 4, 4, 4, 6, 6, 6, 6, 7, 8)
        Case 1: t = Array(0, 1, 1, 1, 2, 2, 4, 4, 4, 5, 5, 5, 8, 9, 9, 10, 10, 11, 13, 14, 16)
        Case 2: t = Array(0, 1, 1, 2, 2, 4, 4, 6, 6, 8, 8, 8, 10, 12, 16, 12, 17, 16, 18, 21, 20)
        Case Else: t = Array(0, 1, 1, 2, 4, 4, 4, 5, 6, 8, 8, 11, 11, 16, 16, 18, 16, 19, 21, 25, 25)
    End Select
    gcNumBlocks = t(LBound(t) + ver)
End Function

Private Function gcRawModules(ByVal ver As Long) As Long
    Dim r As Long, na As Long
    r = ((16 * ver + 128) * ver) + 64
    If ver >= 2 Then
        na = (ver \ 7) + 2
        r = r - (((25 * na - 10) * na) - 55)
        If ver >= 7 Then r = r - 36
    End If
    gcRawModules = r
End Function

Private Function gcDataCodewords(ByVal ver As Long, ByVal ecl As Long) As Long
    gcDataCodewords = (gcRawModules(ver) \ 8) - (gcEccPerBlock(ecl, ver) * gcNumBlocks(ecl, ver))
End Function

'---------- Reed-Solomon ----------
Private Function gcGfMul(ByVal x As Long, ByVal y As Long) As Long
    Dim z As Long, i As Long
    z = 0
    For i = 7 To 0 Step -1
        z = (z * 2) Xor ((z \ 128) * &H11D&)
        If gcBit(y, i) = 1 Then z = z Xor x
    Next i
    gcGfMul = z
End Function

Private Sub gcRsDivisor(ByVal degree As Long, ByRef res() As Long)
    Dim i As Long, j As Long, root As Long
    For i = 0 To degree - 1
        res(i) = 0
    Next i
    res(degree - 1) = 1
    root = 1
    For i = 0 To degree - 1
        For j = 0 To degree - 1
            res(j) = gcGfMul(res(j), root)
            If j + 1 < degree Then res(j) = res(j) Xor res(j + 1)
        Next j
        root = gcGfMul(root, 2)
    Next i
End Sub

Private Sub gcRsRemainder(ByRef data() As Long, ByVal start As Long, ByVal n As Long, _
                          ByRef divisor() As Long, ByVal degree As Long, ByRef res() As Long)
    Dim i As Long, j As Long, factor As Long
    For i = 0 To degree - 1
        res(i) = 0
    Next i
    For i = 0 To n - 1
        factor = data(start + i) Xor res(0)
        For j = 0 To degree - 2
            res(j) = res(j + 1)
        Next j
        res(degree - 1) = 0
        For j = 0 To degree - 1
            res(j) = res(j) Xor gcGfMul(divisor(j), factor)
        Next j
    Next i
End Sub

'---------- bit writer ----------
Private Sub gcPutBits(ByRef buf() As Long, ByRef bitLen As Long, ByVal v As Long, ByVal cnt As Long)
    Dim i As Long
    For i = cnt - 1 To 0 Step -1
        If gcBit(v, i) = 1 Then
            buf(bitLen \ 8) = buf(bitLen \ 8) Or gcShl(1, 7 - (bitLen Mod 8))
        End If
        bitLen = bitLen + 1
    Next i
End Sub

'---------- matrix drawing helpers ----------
Private Sub gcSetFn(ByVal x As Long, ByVal y As Long, ByVal dark As Boolean)
    qrMod(y, x) = dark
    qrFn(y, x) = True
End Sub

Private Sub gcFinder(ByVal cx As Long, ByVal cy As Long)
    Dim dx As Long, dy As Long, dist As Long, xx As Long, yy As Long
    For dy = -4 To 4
        For dx = -4 To 4
            dist = Abs(dx)
            If Abs(dy) > dist Then dist = Abs(dy)
            xx = cx + dx
            yy = cy + dy
            If xx >= 0 And xx < qrSize And yy >= 0 And yy < qrSize Then
                gcSetFn xx, yy, (dist <> 2 And dist <> 4)
            End If
        Next dx
    Next dy
End Sub

Private Sub gcAlign(ByVal cx As Long, ByVal cy As Long)
    Dim dx As Long, dy As Long, dist As Long
    For dy = -2 To 2
        For dx = -2 To 2
            dist = Abs(dx)
            If Abs(dy) > dist Then dist = Abs(dy)
            gcSetFn cx + dx, cy + dy, (dist <> 1)
        Next dx
    Next dy
End Sub

Private Sub gcFormatBits(ByVal ecl As Long, ByVal mask As Long)
    Dim fmt As Long, data As Long, rem2 As Long, bits As Long, i As Long
    Select Case ecl
        Case 0: fmt = 1
        Case 1: fmt = 0
        Case 2: fmt = 3
        Case Else: fmt = 2
    End Select
    data = gcShl(fmt, 3) Or mask
    rem2 = data
    For i = 1 To 10
        rem2 = gcShl(rem2, 1) Xor (gcShr(rem2, 9) * &H537&)
    Next i
    bits = (gcShl(data, 10) Or rem2) Xor &H5412&
    For i = 0 To 5
        gcSetFn 8, i, (gcBit(bits, i) = 1)
    Next i
    gcSetFn 8, 7, (gcBit(bits, 6) = 1)
    gcSetFn 8, 8, (gcBit(bits, 7) = 1)
    gcSetFn 7, 8, (gcBit(bits, 8) = 1)
    For i = 9 To 14
        gcSetFn 14 - i, 8, (gcBit(bits, i) = 1)
    Next i
    For i = 0 To 7
        gcSetFn qrSize - 1 - i, 8, (gcBit(bits, i) = 1)
    Next i
    For i = 8 To 14
        gcSetFn 8, qrSize - 15 + i, (gcBit(bits, i) = 1)
    Next i
    gcSetFn 8, qrSize - 8, True
End Sub

Private Sub gcVersionBits(ByVal ver As Long)
    Dim rem2 As Long, bits As Long, i As Long, a As Long, b As Long
    If ver < 7 Then Exit Sub
    rem2 = ver
    For i = 1 To 12
        rem2 = gcShl(rem2, 1) Xor (gcShr(rem2, 11) * &H1F25&)
    Next i
    bits = gcShl(ver, 12) Or rem2
    For i = 0 To 17
        a = qrSize - 11 + (i Mod 3)
        b = i \ 3
        gcSetFn a, b, (gcBit(bits, i) = 1)
        gcSetFn b, a, (gcBit(bits, i) = 1)
    Next i
End Sub

Private Sub gcFunctionPatterns(ByVal ver As Long, ByVal ecl As Long)
    Dim i As Long, j As Long, na As Long, stepAlign As Long
    Dim pos() As Long
    For i = 0 To qrSize - 1
        gcSetFn 6, i, (i Mod 2 = 0)
        gcSetFn i, 6, (i Mod 2 = 0)
    Next i
    gcFinder 3, 3
    gcFinder qrSize - 4, 3
    gcFinder 3, qrSize - 4
    If ver > 1 Then
        na = (ver \ 7) + 2
        If ver = 32 Then
            stepAlign = 26
        Else
            stepAlign = (((ver * 4) + (na * 2) + 1) \ ((na * 2) - 2)) * 2
        End If
        ReDim pos(0 To na - 1)
        pos(0) = 6
        For i = 1 To na - 1
            pos(i) = qrSize - 7 - ((na - 1 - i) * stepAlign)
        Next i
        For i = 0 To na - 1
            For j = 0 To na - 1
                If Not ((i = 0 And j = 0) Or (i = 0 And j = na - 1) Or (i = na - 1 And j = 0)) Then
                    gcAlign pos(i), pos(j)
                End If
            Next j
        Next i
    End If
    gcFormatBits ecl, 0
    gcVersionBits ver
End Sub

Private Sub gcPlaceCodewords(ByRef cw() As Long, ByVal total As Long)
    Dim i As Long, pr As Long, rc As Long, vert As Long, j As Long, x As Long, y As Long, upward As Boolean
    i = 0
    pr = qrSize - 1
    Do While pr >= 1
        rc = pr
        If rc <= 6 Then rc = rc - 1
        For vert = 0 To qrSize - 1
            For j = 0 To 1
                x = rc - j
                upward = (((rc + 1) And 2) = 0)
                If upward Then y = qrSize - 1 - vert Else y = vert
                If Not qrFn(y, x) And i < total * 8 Then
                    qrMod(y, x) = (gcBit(cw(i \ 8), 7 - (i Mod 8)) = 1)
                    i = i + 1
                End If
            Next j
        Next vert
        pr = pr - 2
    Loop
End Sub

Private Sub gcApplyMask(ByVal mask As Long)
    Dim x As Long, y As Long, inv As Boolean
    For y = 0 To qrSize - 1
        For x = 0 To qrSize - 1
            Select Case mask
                Case 0: inv = (((x + y) Mod 2) = 0)
                Case 1: inv = ((y Mod 2) = 0)
                Case 2: inv = ((x Mod 3) = 0)
                Case 3: inv = (((x + y) Mod 3) = 0)
                Case 4: inv = ((((x \ 3) + (y \ 2)) Mod 2) = 0)
                Case 5: inv = ((((x * y) Mod 2) + ((x * y) Mod 3)) = 0)
                Case 6: inv = (((((x * y) Mod 2) + ((x * y) Mod 3)) Mod 2) = 0)
                Case Else: inv = (((((x + y) Mod 2) + ((x * y) Mod 3)) Mod 2) = 0)
            End Select
            If inv And Not qrFn(y, x) Then qrMod(y, x) = Not qrMod(y, x)
        Next x
    Next y
End Sub

'---------- penalty ----------
Private Sub gcFinderAddHistory(ByVal runLen As Long, ByRef rh() As Long)
    Dim i As Long
    If rh(0) = 0 Then runLen = runLen + qrSize
    For i = 6 To 1 Step -1
        rh(i) = rh(i - 1)
    Next i
    rh(0) = runLen
End Sub

Private Function gcFinderCount(ByRef rh() As Long) As Long
    Dim n As Long, core As Boolean, r As Long
    n = rh(1)
    core = (n > 0 And rh(2) = n And rh(3) = n * 3 And rh(4) = n And rh(5) = n)
    r = 0
    If core And rh(0) >= n * 4 And rh(6) >= n Then r = r + 1
    If core And rh(6) >= n * 4 And rh(0) >= n Then r = r + 1
    gcFinderCount = r
End Function

Private Function gcFinderTerminate(ByVal runColor As Boolean, ByVal runLen As Long, ByRef rh() As Long) As Long
    If runColor Then
        gcFinderAddHistory runLen, rh
        runLen = 0
    End If
    runLen = runLen + qrSize
    gcFinderAddHistory runLen, rh
    gcFinderTerminate = gcFinderCount(rh)
End Function

Private Function gcPenalty() As Long
    Dim res As Long, x As Long, y As Long, i As Long
    Dim runColor As Boolean, cur As Boolean, runLen As Long
    Dim rh(0 To 6) As Long
    Dim dark As Long, tot As Long, k As Long

    res = 0
    For y = 0 To qrSize - 1
        runColor = False: runLen = 0
        For i = 0 To 6
            rh(i) = 0
        Next i
        For x = 0 To qrSize - 1
            cur = qrMod(y, x)
            If cur = runColor Then
                runLen = runLen + 1
                If runLen = 5 Then
                    res = res + 3
                ElseIf runLen > 5 Then
                    res = res + 1
                End If
            Else
                gcFinderAddHistory runLen, rh
                If Not runColor Then res = res + (gcFinderCount(rh) * 40)
                runColor = cur
                runLen = 1
            End If
        Next x
        res = res + (gcFinderTerminate(runColor, runLen, rh) * 40)
    Next y

    For x = 0 To qrSize - 1
        runColor = False: runLen = 0
        For i = 0 To 6
            rh(i) = 0
        Next i
        For y = 0 To qrSize - 1
            cur = qrMod(y, x)
            If cur = runColor Then
                runLen = runLen + 1
                If runLen = 5 Then
                    res = res + 3
                ElseIf runLen > 5 Then
                    res = res + 1
                End If
            Else
                gcFinderAddHistory runLen, rh
                If Not runColor Then res = res + (gcFinderCount(rh) * 40)
                runColor = cur
                runLen = 1
            End If
        Next y
        res = res + (gcFinderTerminate(runColor, runLen, rh) * 40)
    Next x

    For y = 0 To qrSize - 2
        For x = 0 To qrSize - 2
            cur = qrMod(y, x)
            If cur = qrMod(y, x + 1) And cur = qrMod(y + 1, x) And cur = qrMod(y + 1, x + 1) Then res = res + 3
        Next x
    Next y

    dark = 0
    For y = 0 To qrSize - 1
        For x = 0 To qrSize - 1
            If qrMod(y, x) Then dark = dark + 1
        Next x
    Next y
    tot = qrSize * qrSize
    k = (((Abs(dark * 20 - tot * 10) + tot - 1) \ tot) - 1)
    res = res + (k * 10)
    gcPenalty = res
End Function

'---------- build the whole QR matrix (byte mode, versions 1-10) ----------
' Returns False if the text is too long.  Result is in qrMod(y, x), size qrSize.
Private Function gcBuildQr(ByVal txt As String, ByVal ecl As Long, ByVal forceMask As Long) As Boolean
    Dim b() As Byte, nb As Long
    Dim ver As Long, ccBits As Long, capBits As Long, needBits As Long
    Dim buf() As Long, bitLen As Long, term As Long, pad As Long, dataCw As Long
    Dim i As Long, j As Long, k As Long
    Dim nBlocks As Long, eccLen As Long, raw As Long, nShort As Long, shortLen As Long
    Dim dl As Long, startAt As Long
    Dim divisor() As Long, ecc() As Long, eccAll() As Long, blkStart() As Long, blkLen() As Long
    Dim cw() As Long, outN As Long
    Dim best As Long, bestPen As Long, pen As Long, m As Long

    gcBuildQr = False
    nb = gcUtf8(txt, b)
    If nb = 0 Then Exit Function

    For ver = 1 To 20
        capBits = gcDataCodewords(ver, ecl) * 8
        If ver <= 9 Then ccBits = 8 Else ccBits = 16
        needBits = 4 + ccBits + (8 * nb)
        If needBits <= capBits Then Exit For
    Next ver
    If ver > 20 Then Exit Function

    dataCw = gcDataCodewords(ver, ecl)
    ReDim buf(0 To dataCw - 1)
    bitLen = 0
    gcPutBits buf, bitLen, 4, 4
    gcPutBits buf, bitLen, nb, ccBits
    For i = 0 To nb - 1
        gcPutBits buf, bitLen, b(i), 8
    Next i
    term = capBits - bitLen
    If term > 4 Then term = 4
    gcPutBits buf, bitLen, 0, term
    bitLen = ((bitLen + 7) \ 8) * 8
    pad = &HEC&
    Do While bitLen < capBits
        gcPutBits buf, bitLen, pad, 8
        If pad = &HEC& Then pad = &H11& Else pad = &HEC&
    Loop

    nBlocks = gcNumBlocks(ecl, ver)
    eccLen = gcEccPerBlock(ecl, ver)
    raw = gcRawModules(ver) \ 8
    nShort = nBlocks - (raw Mod nBlocks)
    shortLen = raw \ nBlocks

    ReDim divisor(0 To eccLen - 1)
    ReDim ecc(0 To eccLen - 1)
    ReDim eccAll(0 To nBlocks - 1, 0 To eccLen - 1)
    ReDim blkStart(0 To nBlocks - 1)
    ReDim blkLen(0 To nBlocks - 1)
    gcRsDivisor eccLen, divisor

    startAt = 0
    For i = 0 To nBlocks - 1
        dl = shortLen - eccLen
        If i >= nShort Then dl = dl + 1
        blkStart(i) = startAt
        blkLen(i) = dl
        gcRsRemainder buf, startAt, dl, divisor, eccLen, ecc
        For j = 0 To eccLen - 1
            eccAll(i, j) = ecc(j)
        Next j
        startAt = startAt + dl
    Next i

    ReDim cw(0 To raw - 1)
    outN = 0
    For k = 0 To shortLen - eccLen
        For i = 0 To nBlocks - 1
            If k < blkLen(i) Then
                cw(outN) = buf(blkStart(i) + k)
                outN = outN + 1
            End If
        Next i
    Next k
    For j = 0 To eccLen - 1
        For i = 0 To nBlocks - 1
            cw(outN) = eccAll(i, j)
            outN = outN + 1
        Next i
    Next j

    qrSize = (ver * 4) + 17
    Dim x As Long, y As Long
    For y = 0 To 96
        For x = 0 To 96
            qrMod(y, x) = False
            qrFn(y, x) = False
        Next x
    Next y
    gcFunctionPatterns ver, ecl
    gcPlaceCodewords cw, raw

    If forceMask >= 0 Then
        best = forceMask
    Else
        best = 0
        bestPen = 2147483647
        For m = 0 To 7
            gcApplyMask m
            gcFormatBits ecl, m
            pen = gcPenalty()
            If pen < bestPen Then
                bestPen = pen
                best = m
            End If
            gcApplyMask m
        Next m
    End If
    gcApplyMask best
    gcFormatBits ecl, best
    gcBuildQr = True
End Function

'==================== CODE 128 ====================
' Fills w() with alternating bar/space widths (bar first). Returns number of entries
' (0 = text cannot be encoded). totalMods = total width in modules.
' Uses code set C for all-digit even-length text, otherwise code set B (ASCII 32-126).
Private Function gcC128Build(ByVal txt As String, ByRef w() As Long, ByRef totalMods As Long) As Long
    Dim t As String, vals() As Long, nv As Long, i As Long, c As Long
    Dim allDigits As Boolean, chk As Long, n As Long, p As String, k As Long

    gcC128Build = 0
    totalMods = 0
    If Len(txt) = 0 Then Exit Function

    allDigits = True
    For i = 1 To Len(txt)
        c = AscW(Mid$(txt, i, 1))
        If c < 32 Or c > 126 Then Exit Function
        If c < 48 Or c > 57 Then allDigits = False
    Next i

    ReDim vals(0 To Len(txt) + 3)
    nv = 0
    If allDigits And ((Len(txt) Mod 2) = 0) Then
        vals(0) = 105
        nv = 1
        For i = 1 To Len(txt) Step 2
            vals(nv) = CLng(Mid$(txt, i, 2))
            nv = nv + 1
        Next i
    Else
        vals(0) = 104
        nv = 1
        For i = 1 To Len(txt)
            vals(nv) = AscW(Mid$(txt, i, 1)) - 32
            nv = nv + 1
        Next i
    End If
    chk = vals(0)
    For i = 1 To nv - 1
        chk = chk + (vals(i) * i)
    Next i
    vals(nv) = chk Mod 103
    nv = nv + 1
    vals(nv) = 106
    nv = nv + 1

    t = "212222222122222221121223121322131222122213122312132212221213221312231212"
    t = t & "112232122132122231113222123122123221223211221132221231213212223112312131"
    t = t & "311222321122321221312212322112322211212123212321232121111323131123131321"
    t = t & "112313132113132311211313231113231311112133112331132131113123113321133121"
    t = t & "313121211331231131213113213311213131311123311321331121312113312311332111"
    t = t & "314111221411431111111224111422121124121421141122141221112214112412122114"
    t = t & "122411142112142211241211221114413111241112134111111242121142121241114212"
    t = t & "124112124211411212421112421211212141214121412121111143111341131141114113"
    t = t & "114311411113411311113141114131311141411131211412211214211232"

    ReDim w(0 To ((nv - 1) * 6) + 6)
    n = 0
    totalMods = 0
    For i = 0 To nv - 1
        If vals(i) = 106 Then
            p = "2331112"
        Else
            p = Mid$(t, (vals(i) * 6) + 1, 6)
        End If
        For k = 1 To Len(p)
            w(n) = CLng(Mid$(p, k, 1))
            totalMods = totalMods + w(n)
            n = n + 1
        Next k
    Next i
    gcC128Build = n
End Function

'---------- multi-column content: "<11 ID No.> | <1 NAME>" -> real values ----------
' Field tokens look like <index header>; only the leading number matters.
' Anything else typed in the content box is kept as normal text. "\n" = new line.
Private Function gcComposeValue(ByVal tpl As String, ByRef items() As String) As String
    Dim res As String, p As Long, q As Long, q2 As Long, e As Long, inner As String, idx As Long

    res = ""
    p = 1
    Do While p <= Len(tpl)
        q = InStr(p, tpl, "<")
        If q = 0 Then
            res = res & Mid$(tpl, p)
            Exit Do
        End If
        res = res & Mid$(tpl, p, q - p)
        e = InStr(q + 1, tpl, ">")
        If e = 0 Then
            res = res & Mid$(tpl, q)
            Exit Do
        End If
        q2 = InStr(q + 1, tpl, "<")
        If q2 > 0 And q2 < e Then
            res = res & "<"
            p = q + 1
        Else
            inner = Mid$(tpl, q + 1, e - q - 1)
            If Len(inner) > 0 And inner Like "#*" Then
                idx = Val(inner)
                If idx >= LBound(items) And idx <= UBound(items) Then res = res & Trim$(items(idx))
            Else
                res = res & Mid$(tpl, q, e - q + 1)
            End If
            p = e + 1
        End If
    Loop
    res = Replace(res, "\n", Chr$(10))
    gcComposeValue = res
End Function

Private Function gcSepText(ByVal nm As String) As String
    Select Case nm
        Case "|":        gcSepText = " | "
        Case ",":        gcSepText = ", "
        Case ";":        gcSepText = "; "
        Case "-":        gcSepText = " - "
        Case "Space":    gcSepText = " "
        Case "New line": gcSepText = "\n"
        Case Else:       gcSepText = ""
    End Select
End Function


'=====================================================================
'  DRAWING INTO CORELDRAW  (vector, one object per card)
'=====================================================================
Private Sub gcAddRect(ByVal crv As Curve, ByVal x1 As Double, ByVal y1 As Double, _
                      ByVal x2 As Double, ByVal y2 As Double)
    Dim sp As SubPath
    Set sp = crv.CreateSubPath(x1, y1)
    sp.AppendLineSegment x2, y1
    sp.AppendLineSegment x2, y2
    sp.AppendLineSegment x1, y2
    sp.Closed = True
End Sub

' cmykK: 100 = black, 0 = white
Private Function gcMakeShape(ByVal crv As Curve, ByVal cmykK As Long) As Shape
    Dim s As Shape
    Set s = ActiveLayer.CreateCurve(crv)
    s.Fill.ApplyUniformFill CreateCMYKColor(0, 0, 0, cmykK)
    s.Outline.SetNoOutline
    Set gcMakeShape = s
End Function

' Draws the code centred at (cx, cy) inside a region of regW x regH (all in INCHES)
Private Function gcAddCode(ByVal value As String, ByVal codeType As String, _
                           ByVal cx As Double, ByVal cy As Double, _
                           ByVal regW As Double, ByVal regH As Double, _
                           ByVal showText As Boolean, ByVal grpName As String) As Boolean
    Dim oldUnit As cdrUnit
    Dim crvBg As Curve, crvBars As Curve
    Dim shpBg As Shape, shpBars As Shape, shpTxt As Shape, grp As Shape
    Dim sr As New ShapeRange
    Dim textH As Double, boxW As Double, boxH As Double, m As Double
    Dim side As Double, availH As Double, fs As Double, xPos As Double
    Dim x As Long, y As Long, x0 As Long, i As Long, q As Long
    Dim w() As Long, nW As Long, totalMods As Long

    gcAddCode = False
    oldUnit = ActiveDocument.Unit
    ActiveDocument.Unit = cdrInch
    On Error GoTo EH

    textH = 0
    If showText Then textH = regH * 0.2

    Set crvBars = CreateCurve(ActiveDocument)

    If codeType = "Code 128" Then
        nW = gcC128Build(value, w, totalMods)
        If nW = 0 Then
            gcLastError = vbCrLf & "Code 128 cannot encode this value (only letters, digits and normal symbols; no Malayalam/accented letters)."
            GoTo Done
        End If
        q = GC_C128_QUIET
        m = regW / (totalMods + (2 * q))
        boxW = regW
        boxH = regH
        xPos = q * m
        For i = 0 To nW - 1
            If (i Mod 2) = 0 Then gcAddRect crvBars, xPos, textH, xPos + (w(i) * m), boxH
            xPos = xPos + (w(i) * m)
        Next i
    Else
        If Not gcBuildQr(value, GC_QR_ECC, -1) Then
            gcLastError = vbCrLf & "The value is too long for the QR code (limit is about 600 characters)."
            GoTo Done
        End If
        q = GC_QR_QUIET
        availH = regH - textH
        side = regW
        If availH < side Then side = availH
        m = side / (qrSize + (2 * q))
        boxW = side
        boxH = side + textH
        For y = 0 To qrSize - 1
            x = 0
            Do While x < qrSize
                If qrMod(y, x) Then
                    x0 = x
                    Do While x < qrSize
                        If Not qrMod(y, x) Then Exit Do
                        x = x + 1
                    Loop
                    gcAddRect crvBars, (x0 + q) * m, textH + ((qrSize - 1 - y + q) * m), _
                                       (x + q) * m, textH + ((qrSize - y + q) * m)
                Else
                    x = x + 1
                End If
            Loop
        Next y
    End If

    If GC_WHITE_BG Then
        Set crvBg = CreateCurve(ActiveDocument)
        gcAddRect crvBg, 0, 0, boxW, boxH
        Set shpBg = gcMakeShape(crvBg, 0)
        sr.Add shpBg
    End If

    Set shpBars = gcMakeShape(crvBars, 100)
    sr.Add shpBars

    If showText Then
        fs = textH * 72 * 0.75
        If fs < 4 Then fs = 4
        Set shpTxt = ActiveLayer.CreateArtisticText(0, 0, Replace(Replace(value, Chr$(13), " "), Chr$(10), " "), , , GC_FONT, fs)
        shpTxt.Text.Story.Fill.ApplyUniformFill CreateCMYKColor(0, 0, 0, 100)
        If shpTxt.SizeWidth > (boxW * 0.96) Then
            shpTxt.SetSize boxW * 0.96, shpTxt.SizeHeight * (boxW * 0.96) / shpTxt.SizeWidth
        End If
        shpTxt.CenterX = boxW / 2
        shpTxt.CenterY = textH / 2
        sr.Add shpTxt
    End If

    Set grp = sr.Group
    grp.CenterX = cx
    grp.CenterY = cy
    grp.Name = grpName
    gcAddCode = True

Done:
    ActiveDocument.Unit = oldUnit
    Exit Function
EH:
    gcLastError = vbCrLf & "Reason: " & Err.Description
    gcAddCode = False
    Resume Done
End Function

'=====================================================================
'  WINDOW LAYOUT (header, navigation rail, pages, footer)  +  QR / BARCODE PANELS
'=====================================================================
Private Function gcFindFrame(ByVal cap As String) As Object
    Dim c As Object
    For Each c In Me.Controls
        If TypeName(c) = "Frame" Then
            If c.Caption = cap Then
                Set gcFindFrame = c
                Exit Function
            End If
        End If
    Next c
    Set gcFindFrame = Nothing
End Function

Private Function gcFindCtl(ByVal nm As String) As Object
    Dim c As Object
    On Error Resume Next
    Set c = Me.Controls(nm)
    On Error GoTo 0
    Set gcFindCtl = c
End Function

Private Function gcChildOf(ByVal c As Object, ByVal parentName As String) As Boolean
    Dim p As Object
    gcChildOf = False
    On Error Resume Next
    Set p = c.Parent
    If Not p Is Nothing Then gcChildOf = (p.Name = parentName)
    On Error GoTo 0
End Function

Private Function gcMax(ByVal a As Single, ByVal b As Single) As Single
    If a > b Then gcMax = a Else gcMax = b
End Function

Private Function gcCap(ByVal c As Object) As String
    gcCap = ""
    On Error Resume Next
    If TypeName(c) = "Frame" Then gcCap = c.Caption
End Function

'---------------------------------------------------------------------
' Modern window: header, left navigation rail, one page at a time, footer.
'   Cards | QR / Barcode | PowerClip | Arrange | Text / Group
' The controls of the old designer are only moved and shown/hidden here.
'---------------------------------------------------------------------
Private Sub gcBuildUi()
    Dim mf As Object, c As Object, fb As Object, fp As Object, dl As Object, st As Object, f As Object
    Dim fChg As Object, fMov As Object, fGrp As Object
    Dim i As Long, nm As String, iconNames As String, vn As Variant, grpNames As Variant
    Dim haveCards As Boolean, haveIcons As Boolean, haveTools As Boolean, haveGrp As Boolean
    Dim cl As Single, ct As Single, cr As Single, cb As Single
    Dim il As Single, it As Single, ir As Single, ib As Single
    Dim tl As Single, tt As Single, tr As Single, tb As Single
    Dim gl As Single, gt As Single, gr As Single, gb As Single
    Dim batchW As Single, batchH As Single, iconW As Single, iconH As Single
    Dim toolsW As Single, toolsH As Single, grpW As Single, grpH As Single
    Dim chgW As Single, chgH As Single, movW As Single, movH As Single, gbyW As Single, gbyH As Single
    Dim stageW As Single, dlH As Single, contentH As Single, h As Single
    Dim padG As Single, railW As Single, headerH As Single, footerH As Single
    Dim formW As Single, formH As Single, y0 As Single, addCards As Boolean

    On Error GoTo EH
    gcUiReady = False
    padG = 8: railW = 104: headerH = 36: footerH = 48

    Set mf = gcFindFrame("Card Maker")
    If mf Is Nothing Then Exit Sub                 ' main frame not found: keep the old window
    Set gcMf = mf
    mf.Caption = ""
    mf.BorderStyle = 0                             ' flat, no caption

    For i = 0 To 4
        Set gcTab(i) = New Collection
    Next i
    iconNames = "|FitImgtoPCframe|FitTextToFrame|Stretch_toFit|DeleteDatainEveryPage|RotateObj|"
    grpNames = Array("GrpMinX", "GrpMaxX", "GrpMinY", "GrpMaxY", "GrpGetBoundingBox", "GrpCmdGroupShapes")

    '---- sort the children of the main frame into pages ----------------------------------------
    For Each c In Me.Controls
        If gcChildOf(c, mf.Name) Then
            nm = c.Name
            addCards = False
            If InStr(iconNames, "|" & nm & "|") > 0 Then
                gcTab(2).Add c                                  ' PowerClip page: quick buttons
                If Not haveIcons Then
                    il = c.Left: it = c.Top: ir = c.Left + c.Width: ib = c.Top + c.Height
                    haveIcons = True
                Else
                    If c.Left < il Then il = c.Left
                    If c.Top < it Then it = c.Top
                    If c.Left + c.Width > ir Then ir = c.Left + c.Width
                    If c.Top + c.Height > ib Then ib = c.Top + c.Height
                End If
            ElseIf gcCap(c) = "Batch PowerClip" Then
                Set fb = c
                gcTab(2).Add c
            ElseIf gcCap(c) = "Bring/Send" Or gcCap(c) = "Resize" Then
                gcTab(3).Add c                                  ' Arrange page
                If Not haveTools Then
                    tl = c.Left: tt = c.Top: tr = c.Left + c.Width: tb = c.Top + c.Height
                    haveTools = True
                Else
                    If c.Left < tl Then tl = c.Left
                    If c.Top < tt Then tt = c.Top
                    If c.Left + c.Width > tr Then tr = c.Left + c.Width
                    If c.Top + c.Height > tb Then tb = c.Top + c.Height
                End If
            ElseIf nm = "createCards" Then
                c.Visible = False                               ' replaced by the Generate button in the footer
            ElseIf nm = "Status" Then
                c.Left = -3000                                  ' old banner: the footer shows the status now
            ElseIf nm = "folderpath" Or nm = "DummyList" Then
                If c.Visible Then gcTab(0).Add c
            ElseIf c.Visible Then
                addCards = True
            End If
            If addCards Then
                gcTab(0).Add c                                  ' Cards page
                If Not haveCards Then
                    cl = c.Left: ct = c.Top: cr = c.Left + c.Width: cb = c.Top + c.Height
                    haveCards = True
                Else
                    If c.Left < cl Then cl = c.Left
                    If c.Top < ct Then ct = c.Top
                    If c.Left + c.Width > cr Then cr = c.Left + c.Width
                    If c.Top + c.Height > cb Then cb = c.Top + c.Height
                End If
            End If
        End If
    Next c
    If Not haveCards Then Exit Sub

    '---- sizes of everything -------------------------------------------------------------------------
    If Not fb Is Nothing Then batchW = fb.Width: batchH = fb.Height
    If haveIcons Then iconW = ir - il: iconH = ib - it
    If haveTools Then toolsW = tr - tl: toolsH = tb - tt
    Set fChg = gcFindFrame("Change Case")
    Set fMov = gcFindFrame("Move Object")
    Set fGrp = gcFindFrame("Group By Shape No.")
    If Not fChg Is Nothing Then chgW = fChg.Width: chgH = fChg.Height
    If Not fMov Is Nothing Then movW = fMov.Width: movH = fMov.Height
    If Not fGrp Is Nothing Then gbyW = fGrp.Width: gbyH = fGrp.Height
    For Each vn In grpNames
        Set c = gcFindCtl(CStr(vn))
        If Not c Is Nothing Then
            If Not haveGrp Then
                gl = c.Left: gt = c.Top: gr = c.Left + c.Width: gb = c.Top + c.Height
                haveGrp = True
            Else
                If c.Left < gl Then gl = c.Left
                If c.Top < gt Then gt = c.Top
                If c.Left + c.Width > gr Then gr = c.Left + c.Width
                If c.Top + c.Height > gb Then gb = c.Top + c.Height
            End If
        End If
    Next vn
    If haveGrp Then grpW = gr - gl: grpH = gb - gt

    Set dl = gcFindCtl("DummyList")
    dlH = 0
    If Not dl Is Nothing Then dlH = dl.Height

    stageW = 240
    For Each vn In Array(cr - cl, batchW, iconW, toolsW, grpW, chgW, movW, gbyW)
        If vn > stageW Then stageW = vn
    Next vn

    contentH = cb - ct + 4 + dlH                       ' Cards page
    If gcQrBlockHeight() > contentH Then contentH = gcQrBlockHeight()
    h = batchH + 8 + iconH
    If h > contentH Then contentH = h
    h = toolsH + 6 + movH
    If h > contentH Then contentH = h
    h = chgH + 8 + gbyH + 8 + grpH
    If h > contentH Then contentH = h

    '---- arrange the pages inside the frame -------------------------------------------------------
    If Not fb Is Nothing Then                          ' PowerClip page
        fb.Left = cl: fb.Top = ct
    End If
    If haveIcons Then
        For Each c In gcTab(2)
            If InStr(iconNames, "|" & c.Name & "|") > 0 Then
                c.Left = c.Left + (cl - il)
                c.Top = c.Top + ((ct + batchH + 8) - it)
            End If
        Next c
    End If
    If haveTools Then                                  ' Arrange page
        For Each c In gcTab(3)
            c.Left = c.Left + (cl - tl)
            c.Top = c.Top + (ct - tt)
        Next c
    End If
    Set fp = gcFindCtl("folderpath")
    If Not fp Is Nothing Then
        If cr - fp.Left > 40 Then fp.Width = cr - fp.Left
    End If
    If Not dl Is Nothing Then
        dl.Left = cl: dl.Top = cb + 4: dl.Width = stageW
    End If

    mf.Width = cl + stageW + padG
    gcFrameH(0) = ct + contentH + padG                 ' Cards, QR, PowerClip
    gcFrameH(1) = gcFrameH(0)
    gcFrameH(2) = gcFrameH(0)
    gcFrameH(3) = ct + toolsH + padG                   ' Arrange (short frame, Move Object below it)
    gcFrameH(4) = gcFrameH(0)
    mf.Left = railW + padG
    mf.Top = headerH + padG
    mf.Height = gcFrameH(0)
    mf.BackColor = &HFFFFFF

    gcBuildQrTab cl, ct, stageW                        ' QR page, inside the frame

    '---- top-level frames (not inside the main frame) --------------------------------------------
    If Not fMov Is Nothing Then                        ' Arrange page, under the short frame
        fMov.Left = mf.Left
        fMov.Top = mf.Top + gcFrameH(3) + 6
        gcTab(3).Add fMov
    End If
    y0 = mf.Top                                       ' Text / Group page
    If Not fChg Is Nothing Then
        fChg.Left = mf.Left: fChg.Top = y0
        gcTab(4).Add fChg
        y0 = y0 + chgH + 8
    End If
    If Not fGrp Is Nothing Then
        fGrp.Left = mf.Left: fGrp.Top = y0
        gcTab(4).Add fGrp
        y0 = y0 + gbyH + 8
    End If
    For Each vn In grpNames
        Set c = gcFindCtl(CStr(vn))
        If Not c Is Nothing Then
            c.Left = c.Left + (mf.Left - gl)
            c.Top = c.Top + (y0 - gt)
            gcTab(4).Add c
        End If
    Next vn

    '---- the old standalone QR controls of the designer are not used any more ---------------
    For Each vn In Array("gcCheck", "gcType", "gcColumn", "gcMove", "gcGetPos", "gcX", "gcY", "gcW", "gcH", "gcText")
        Set c = gcFindCtl(CStr(vn))
        If Not c Is Nothing Then c.Visible = False
    Next vn

    '---- window size --------------------------------------------------------------------------------------
    formW = railW + padG + mf.Width + padG
    formH = headerH + padG + gcFrameH(0) + padG + footerH
    Me.Width = Me.Width - Me.InsideWidth + formW
    Me.Height = Me.Height - Me.InsideHeight + formH
    Me.BackColor = &HF8F5F3

    gcBuildChrome formW, formH, railW, headerH, footerH, padG
    gcRecolor

    gcInitPanels
    gcBusy = False
    gcUiReady = True
    gcShowPage 0
    Exit Sub
EH:
    gcUiReady = False
    MsgBox "The new window could not be built (" & Err.Description & ")." & vbCrLf & _
           "The tools still work; send me a screenshot of this message.", vbExclamation
End Sub

'---------------------------------------------------------------------
'  header, navigation rail, footer (made by code, outside the main frame)
'---------------------------------------------------------------------
Private Function gcLbl(ByVal nm As String, ByVal l As Single, ByVal t As Single, ByVal w As Single, _
                       ByVal hgt As Single, ByVal cap As String, ByVal fore As Long, ByVal back As Long, _
                       ByVal sz As Single, ByVal bold As Boolean, ByVal algn As Long) As Object
    Dim c As Object
    Set c = Me.Controls.Add("Forms.Label.1", nm, True)
    c.Left = l: c.Top = t: c.Width = w: c.Height = hgt
    c.Caption = cap
    c.ForeColor = fore
    If back < 0 Then
        c.BackStyle = 0
    Else
        c.BackStyle = 1
        c.BackColor = back
    End If
    c.Font.Name = "Tahoma"
    c.Font.Size = sz
    c.Font.Bold = bold
    c.TextAlign = algn
    c.BorderStyle = 0
    Set gcLbl = c
End Function

Private Sub gcBuildChrome(ByVal formW As Single, ByVal formH As Single, ByVal railW As Single, _
                          ByVal headerH As Single, ByVal footerH As Single, ByVal padG As Single)
    Dim c As Object, i As Long, caps As Variant, footTop As Single, tw As Single

    ' header
    gcLbl "gcHdrBar", 0, 0, formW, headerH, "", GC_NAVY, GC_NAVY, 8, False, 1
    gcLbl "gcHdrTitle", 12, 9, 90, 20, "Card Maker", &HFFFFFF, -1, 12, True, 1
    gcLbl "gcHdrSub", 104, 13, formW - 104 - 90, 16, "Printo Cards and Technologies", GC_LIGHT, -1, 8, False, 1
    Set gcChip = gcLbl("gcChip", formW - 76, 9, 64, 18, "Ready", &HFFFFFF, GC_TEAL, 8, True, 2)

    ' navigation rail
    gcLbl "gcRail", 0, headerH, railW, formH - headerH - footerH, "", GC_NAVY, GC_RAIL, 8, False, 1
    caps = Array("Cards", "QR / Barcode", "PowerClip", "Arrange", "Text / Group")
    For i = 1 To 5
        Set c = gcLbl("gcNav" & i, 0, headerH + 10 + (i - 1) * 34, railW, 32, vbCr & "   " & caps(i - 1), GC_NAVY, GC_RAIL, 9, False, 1)
        Select Case i
            Case 1: Set nv1 = c
            Case 2: Set nv2 = c
            Case 3: Set nv3 = c
            Case 4: Set nv4 = c
            Case 5: Set nv5 = c
        End Select
    Next i

    ' footer
    footTop = formH - footerH
    gcLbl "gcFtBar", 0, footTop, formW, footerH, "", GC_NAVY, &HFFFFFF, 8, False, 1
    gcLbl "gcFtLine", 0, footTop, formW, 1, "", GC_NAVY, &HD1C9C2, 8, False, 1
    Set goBtn = gcLbl("gcGo", padG, footTop + 9, 124, 30, vbCr & "Generate cards", &HFFFFFF, GC_TEAL, 10, True, 2)
    tw = formW - padG - 136 - padG
    Set gcTrack = gcLbl("gcPgTrack", padG + 136, footTop + 14, tw, 8, "", GC_NAVY, GC_RAIL, 8, False, 1)
    Set gcFill = gcLbl("gcPgFill", padG + 136, footTop + 14, 0.5, 8, "", GC_NAVY, GC_TEAL, 8, False, 1)
    Set gcFtText = gcLbl("gcFtText", padG + 136, footTop + 27, tw, 16, "Ready", GC_MUTED, -1, 8, False, 1)
End Sub

' every default-grey frame / label / option button gets the white card colour
Private Sub gcRecolor()
    Dim c As Object, tn As String
    On Error Resume Next
    For Each c In Me.Controls
        tn = TypeName(c)
        If tn = "Frame" Or tn = "OptionButton" Or tn = "CheckBox" Or tn = "Label" Then
            If Left$(c.Name, 2) <> "gc" Then
                If c.BackColor = &H8000000F Then c.BackColor = &HFFFFFF
            End If
        End If
    Next c
End Sub

'---------------------------------------------------------------------
'  pages
'---------------------------------------------------------------------
Private Sub nv1_Click()
    gcShowPage 0
End Sub
Private Sub nv2_Click()
    gcShowPage 1
End Sub
Private Sub nv3_Click()
    gcShowPage 2
End Sub
Private Sub nv4_Click()
    gcShowPage 3
End Sub
Private Sub nv5_Click()
    gcShowPage 4
End Sub
Private Sub goBtn_Click()
    createCards_Click
End Sub

Private Sub gcShowPage(ByVal idx As Long)
    Dim p As Long, c As Object
    For p = 0 To 4
        If p <> idx Then
            For Each c In gcTab(p)
                c.Visible = False
            Next c
        End If
    Next p
    gcMf.Visible = (idx <> 4)
    gcMf.Height = gcFrameH(idx)
    For Each c In gcTab(idx)
        c.Visible = True
    Next c
    If idx = 1 Then gcSyncFields
    For p = 0 To 4
        Set c = gcFindCtl("gcNav" & (p + 1))
        If Not c Is Nothing Then
            If p = idx Then
                c.BackColor = &HFFFFFF
                c.ForeColor = GC_TEAL
                c.Font.Bold = True
            Else
                c.BackColor = GC_RAIL
                c.ForeColor = GC_NAVY
                c.Font.Bold = False
            End If
        End If
    Next p
End Sub

'---------------------------------------------------------------------
'  status in the footer (replaces the old black banner)
'---------------------------------------------------------------------
Private Sub gcSetStatus(ByVal txt As String)
    Dim p As Long, q As Long, a As Double, b As Double
    If Not gcUiReady Then Exit Sub
    If Not gcBusy Then
        gcBusy = True
        gcT0 = Timer
        gcChip.Caption = "Working"
        gcChip.BackColor = &H3C8CE8                  ' orange
    End If
    p = InStr(txt, "Page ")
    q = InStr(txt, " of ")
    If p > 0 And q > p Then
        a = Val(Mid$(txt, p + 5))
        b = Val(Mid$(txt, q + 4))
        If b > 0 Then gcFill.Width = gcTrack.Width * a / b
    End If
    gcFtText.Caption = txt & "   -   " & Format(Timer - gcT0, "0") & " s"
End Sub

Private Sub gcIdle()
    If Not gcUiReady Then Exit Sub
    If gcBusy Then gcFtText.Caption = "Finished in " & Format(Timer - gcT0, "0") & " s"
    gcBusy = False
    gcChip.Caption = "Ready"
    gcChip.BackColor = GC_TEAL
End Sub

Private Sub gcNote(ByVal txt As String)
    If gcUiReady Then gcFtText.Caption = txt
End Sub

'---------------------------------------------------------------------
'  QR / Barcode page (inside the main frame)
'---------------------------------------------------------------------
Private Function gcNewCtl(ByVal progId As String, ByVal nm As String, ByVal l As Single, _
                          ByVal t As Single, ByVal w As Single, ByVal hgt As Single, _
                          Optional ByVal cap As String = "") As Object
    Dim c As Object
    Set c = gcMf.Controls.Add(progId, nm, True)
    c.Left = l: c.Top = t: c.Width = w: c.Height = hgt
    If cap <> "" Then c.Caption = cap
    c.Font.Name = "Tahoma"
    c.Font.Size = 8
    gcTab(1).Add c
    Set gcNewCtl = c
End Function

Private Function gcQrBlockHeight() As Single
    gcQrBlockHeight = 18 + 90 + 8 + 90
End Function

Private Sub gcBuildQrTab(ByVal ox As Single, ByVal oy As Single, ByVal w As Single)
    Dim c As Object
    Set c = gcNewCtl("Forms.Label.1", "gcLblQr", ox, oy, w, 14, "QR / Barcode on the card")
    c.Font.Bold = True
    gcBuildSlot 1, ox, oy + 18, w
    gcBuildSlot 2, ox, oy + 18 + 90 + 8, w
End Sub

Private Sub gcBuildSlot(ByVal slot As Long, ByVal x As Single, ByVal y As Single, ByVal w As Single)
    Dim p As String, c As Object
    p = "gc" & slot

    gcNewCtl "Forms.Label.1", "gcBg" & slot, x, y, w, 90                        ' coloured panel

    ' row 0: on/off, type, show text
    Set c = gcNewCtl("Forms.CheckBox.1", p & "Check", x + 4, y + 4, 80, 16, "Generator " & slot)
    c.Font.Bold = True
    Set c = gcNewCtl("Forms.ComboBox.1", p & "Type", x + 88, y + 3, 84, 18)
    c.Style = 2
    gcNewCtl "Forms.CheckBox.1", p & "Text", x + 176, y + 4, w - 180, 16, "Show text"

    ' row 1: field list + Add
    gcNewCtl "Forms.Label.1", "gcLbF" & slot, x + 4, y + 27, 26, 14, "Field"
    Set c = gcNewCtl("Forms.ComboBox.1", p & "Field", x + 32, y + 24, w - 32 - 4 - 54, 18)
    c.Style = 2
    Set c = gcNewCtl("Forms.CommandButton.1", p & "Add", x + w - 4 - 50, y + 24, 50, 18, "Add")
    If slot = 1 Then Set d1Add = c Else Set d2Add = c

    ' row 2: content
    Set c = gcNewCtl("Forms.TextBox.1", p & "Col", x + 4, y + 46, w - 8, 18)
    c.ControlTipText = "Fields in <> are replaced for every card. Type your own text or separators too, e.g.  <4 ID No.> | <1 NAME>   (\n = new line)"

    ' row 3: position
    Set c = gcNewCtl("Forms.CommandButton.1", p & "Pos", x + 4, y + 68, 60, 18, "Position")
    If slot = 1 Then Set d1Pos = c Else Set d2Pos = c
    gcNewCtl "Forms.Label.1", p & "PosLbl", x + 68, y + 71, w - 72, 14, "Position not set"
End Sub

' items, colours and labels of the QR panels
Private Sub gcInitPanels()
    Dim s As Long, cmb As Object
    For s = 1 To 2
        Set cmb = gcCtl(s, "Type")
        If cmb.ListCount = 0 Then
            cmb.AddItem "QR Code"
            cmb.AddItem "Code 128"
        End If
        If cmb.ListIndex < 0 Then cmb.ListIndex = 0
        gcSetPosLabel s
    Next s
    gcApplyColors
End Sub

' every generator has its own colour
Private Sub gcApplyColors()
    Dim s As Long, tint As Long, tint2 As Long, accent As Long, c As Object, vn As Variant
    For s = 1 To 2
        If s = 1 Then
            tint = &HFFE9D6: tint2 = &HF5CDAA: accent = &HAC611F          ' blue
        Else
            tint = &HD6F5D6: tint2 = &HAAE1AA: accent = &H3C821E          ' green
        End If
        Set c = gcFindCtl("gcBg" & s)
        If Not c Is Nothing Then
            c.BackColor = tint
            c.BorderStyle = 1
            c.BorderColor = accent
        End If
        For Each vn In Array("gc" & s & "Check", "gc" & s & "Text", "gcLbF" & s, "gc" & s & "PosLbl")
            Set c = gcFindCtl(CStr(vn))
            If Not c Is Nothing Then c.BackStyle = 0
        Next vn
        Set c = gcFindCtl("gc" & s & "Check")
        If Not c Is Nothing Then c.ForeColor = accent
        For Each vn In Array("gc" & s & "Add", "gc" & s & "Pos")
            Set c = gcFindCtl(CStr(vn))
            If Not c Is Nothing Then c.BackColor = tint2
        Next vn
    Next s
End Sub

Private Sub gcSetPosLabel(ByVal slot As Long)
    Dim c As Object
    Set c = gcCtl(slot, "PosLbl")
    If gcGeoOk(slot) Then
        c.Caption = "Position set  " & Format(gcGeo(slot, 3), "0.00") & " x " & Format(gcGeo(slot, 4), "0.00") & " in"
        c.ForeColor = &H8000&
    Else
        c.Caption = "Position not set"
        c.ForeColor = &HC0&
    End If
End Sub

'---------------------------------------------------------------------
'  access to the controls of generator 1 / 2
'---------------------------------------------------------------------
Private Function gcCtl(ByVal slot As Long, ByVal part As String) As Object
    Set gcCtl = gcMf.Controls("gc" & slot & part)
End Function

Private Function gcIsOn(ByVal slot As Long) As Boolean
    gcIsOn = False
    If gcUiReady Then gcIsOn = CBool(gcCtl(slot, "Check").Value)
End Function

Private Sub gcSnapshotFields()
    Dim i As Long
    Set gcFields = New Collection
    For i = 0 To ListBox1.ListCount - 1
        gcFields.Add ListBox1.List(i)
    Next i
    gcSyncFields
End Sub

Private Sub gcSyncFields()
    Dim s As Long, i As Long, cmb As Object, prev As Long
    If Not gcUiReady Then Exit Sub
    For s = 1 To 2
        Set cmb = gcCtl(s, "Field")
        prev = cmb.ListIndex
        cmb.Clear
        If Not gcFields Is Nothing Then
            For i = 1 To gcFields.Count
                cmb.AddItem gcFields(i)
            Next i
        End If
        If cmb.ListCount = 0 Then
            For i = 0 To ListBox1.ListCount - 1
                cmb.AddItem ListBox1.List(i)
            Next i
        End If
        If prev >= 0 And prev < cmb.ListCount Then cmb.ListIndex = prev
    Next s
End Sub

'---------------------------------------------------------------------
'  buttons of the QR panels
'---------------------------------------------------------------------
Private Sub d1Add_Click()
    gcDoAdd 1
End Sub
Private Sub d1Pos_Click()
    gcDoPos 1
End Sub
Private Sub d2Add_Click()
    gcDoAdd 2
End Sub
Private Sub d2Pos_Click()
    gcDoPos 2
End Sub

' Adds the chosen data field to the content box, e.g.  <11 ID No.>
' Type separators (| , - space) yourself with the keyboard.
Private Sub gcDoAdd(ByVal slot As Long)
    Dim cmb As Object, tb As Object
    Set cmb = gcCtl(slot, "Field")
    Set tb = gcCtl(slot, "Col")
    If cmb.ListIndex < 0 Then
        MsgBox "Choose a data field in the Field list first (load your TXT file first).", vbExclamation
        Exit Sub
    End If
    tb.Value = tb.Value & "<" & cmb.List(cmb.ListIndex) & ">"
End Sub

' Reads position + size of the selected box shape (inches) and keeps them for Generate Cards
Private Sub gcDoPos(ByVal slot As Long)
    Dim sh As Shape, oldU As cdrUnit
    If Documents.Count = 0 Then Exit Sub
    If ActiveSelectionRange.Count <> 1 Then
        MsgBox "Select ONE shape (the box where the QR/barcode should go) first.", vbExclamation
        Exit Sub
    End If
    Set sh = ActiveShape
    oldU = ActiveDocument.Unit
    ActiveDocument.Unit = cdrInch
    gcGeo(slot, 1) = sh.CenterX
    gcGeo(slot, 2) = sh.CenterY
    gcGeo(slot, 3) = sh.SizeWidth
    gcGeo(slot, 4) = sh.SizeHeight
    ActiveDocument.Unit = oldU
    gcGeoOk(slot) = True
    gcSetPosLabel slot
End Sub

Private Function gcNum(ByVal v As String) As Boolean
    gcNum = (Len(Trim$(v)) > 0 And IsNumeric(v))
End Function

' finds the group made by Generator <slot> on the active page
Private Function gcFindCode(ByVal slot As Long) As Shape
    Dim sh As Shape
    Set gcFindCode = Nothing
    For Each sh In ActivePage.ActiveLayer.Shapes
        If sh.Name = "GeneratedCode" & slot Then
            Set gcFindCode = sh
            Exit Function
        End If
    Next sh
End Function

Function CheckDuplicatePlaceholders() As Boolean

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim s As Shape
    Dim txt As String
    Dim startPos As Integer
    Dim endPos As Integer
    Dim key As String
    Dim duplicates As String

    duplicates = ""

    For Each s In ActiveDocument.ActivePage.Shapes
    
        If s.Type = cdrTextShape Then
        
            txt = s.text.Story
            
            Do While InStr(txt, "<") > 0 And InStr(txt, ">") > 0
            
                startPos = InStr(txt, "<")
                endPos = InStr(startPos, txt, ">")
                
                If endPos = 0 Then Exit Do
                
                key = Mid(txt, startPos, endPos - startPos + 1)
                
                If dict.Exists(key) Then
                
                    If InStr(duplicates, key) = 0 Then
                        duplicates = duplicates & key & vbCrLf
                    End If
                    
                Else
                    dict.Add key, 1
                End If
                
                txt = Mid(txt, endPos + 1)
                
            Loop
            
        End If
        
    Next s

    If duplicates <> "" Then
    
        Dim msg As String
        
        msg = "Duplicate placeholders found:" & vbCrLf & vbCrLf & duplicates & vbCrLf & _
              "Do you want to proceed?"

        If MsgBox(msg, vbExclamation + vbYesNo, "Duplicate Warning") = vbNo Then
            CheckDuplicatePlaceholders = False
        Else
            CheckDuplicatePlaceholders = True
        End If
        
    Else
    
        CheckDuplicatePlaceholders = True
        
    End If

End Function


Private Sub CreateTextFields_Click()
    Dim response As VbMsgBoxResult
    response = MsgBox("This action will insert all text fields. Continue?", vbQuestion + vbOKCancel, "Please confirm...")

    If response = vbCancel Then
        ' User pressed "Cancel," so we exit the sub.
        Exit Sub
    End If

    Dim DataText As String
    Dim str As String

    ' Check if a file has been added to the data list (ListBox1).
    If ListBox1.ListCount = 0 Then
        MsgBox "Please add a CSV data file first.", vbOKOnly, "Database Alert - Select CSV file first"
        Exit Sub
    End If

    For iTlist = 0 To ListBox1.ListCount - 1
        DataText = ListBox1.List(iTlist)
        num = iTlist + 1
        Number = CStr(num)
        tshape = "str" + Number
        Set sh = ActiveDocument.ActivePage.ActiveLayer.CreateArtisticText(-2, 4 - (num / 2), "<" + DataText + ">")
        sh.Fill.UniformColor.CMYKAssign 0, 0, 0, 100
        sh.Outline.SetNoOutline
    Next iTlist
End Sub
Function OpenTextFileAuto(filePath As String) As Object
    Dim objFSO As Object
    Dim fnum As Integer
    Dim b1 As Byte, b2 As Byte
    Dim isUnicode As Boolean

    ' Peek at the first 2 bytes to detect a UTF-16 LE BOM (FF FE)
    fnum = FreeFile
    Open filePath For Binary Access Read As #fnum
    If LOF(fnum) >= 2 Then
        Get #fnum, 1, b1
        Get #fnum, 2, b2
    Else
        b1 = 0: b2 = 0
    End If
    Close #fnum

    isUnicode = (b1 = &HFF And b2 = &HFE)

    Set objFSO = CreateObject("Scripting.FileSystemObject")
    If isUnicode Then
        Set OpenTextFileAuto = objFSO.OpenTextFile(filePath, 1, False, -1) ' True Unicode file
    Else
        Set OpenTextFileAuto = objFSO.OpenTextFile(filePath, 1, False, 0)  ' ANSI/system default
    End If
End Function

Private Sub CSV_Select_Click()
Dim FileName As String
Dim objStream As Object
Dim LineFromFile As String
Dim LineItems() As String
Dim str1 As String
Dim str4 As Integer
Dim i As Integer

ListBox1.Clear
ImgList.value = ""
BarList.value = ""
ImgMove.Enabled = True
BarcodeMove.Enabled = True

    FileName = CorelScriptTools.GetFileBox("TXT Files (*.txt)|*.txt", "Choose CSV file", 0, "*.txt", "", "", "")
    
    If FileName = "" Then
        Exit Sub
    Else
        folderpath.value = FileName
    End If
    
    ' Use the special Unicode function to open the file
    Set objStream = OpenTextFileAuto(FileName)

    If Not objStream.AtEndOfStream Then
        LineFromFile = objStream.ReadLine
        
        ' Strip BOM character if present (Safety check)
        If Len(LineFromFile) > 0 Then
            If AscW(Left(LineFromFile, 1)) = -257 Then
                LineFromFile = Mid(LineFromFile, 2)
            End If
        End If
    End If
    objStream.Close

    LineItems() = Split(LineFromFile, vbTab)
    str4 = UBound(LineItems) - LBound(LineItems) + 1
    
    For i = 0 To (str4 - 1)
        str1 = CStr(i) + " " + LineItems(i)
        ListBox1.AddItem str1
    Next i
    gcSnapshotFields

    Set objStream = Nothing
End Sub
Private Sub DoPowerclip_Click()
    Optimization = True
    EventsEnabled = False
    ActiveDocument.SaveSettings
    ActiveDocument.PreserveSelection = False
    
    ' Configure and display the status label
    gcSetStatus "Processing PowerClip Operation"
    DoEvents ' Ensure initial UI update
    
    On Error Resume Next
    ActiveDocument.Unit = cdrCentimeter

    ' Validate the input
    Dim shapeNumber As Integer
    Dim bitmapNumber As Integer
    shapeNumber = txtShapeNumberPC.value
    bitmapNumber = txtBitmapNumberPC.value

    If shapeNumber <= 0 Or bitmapNumber = 0 Then
        MsgBox "Invalid shape or bitmap number. Exiting macro.", vbExclamation
        ActiveDocument.RestoreSettings
        EventsEnabled = True
        Optimization = False
        gcIdle
        ActiveDocument.EndCommandGroup
        ActiveWindow.Refresh
        Application.Refresh
        Exit Sub
    End If

    ' Group commands for undo/redo efficiency
    ActiveDocument.BeginCommandGroup "PowerClip Operation"
    
    ' Check which option is selected and call the respective subroutine
    If CTopt.value = True Then
        Call AlignAndPowerClip
    ElseIf Copt.value = True Then
        Call PowerClipOnly
    Else
        ' If neither option is selected, display a warning message
        MsgBox "Please select an option before executing.", vbExclamation
    End If
    
    ' Finalize command group
    ActiveDocument.EndCommandGroup
    
    ' Clean up and reset environment
    gcIdle
    ActiveDocument.PreserveSelection = True
    ActiveDocument.RestoreSettings
    EventsEnabled = True
    Optimization = False
    
    ' Refresh the UI
    ActiveWindow.Refresh
    Application.Refresh
End Sub

Private Sub ErrorHandleRefresh_Click()
ActiveDocument.PreserveSelection = True
    ActiveDocument.RestoreSettings
    EventsEnabled = True
    Optimization = False
    ActiveDocument.EndCommandGroup
    ActiveWindow.Refresh
    Application.Refresh
End Sub

Private Sub FitImgtoPCframe_Click()
    Dim s As Shape
    Dim srPowerClipped As ShapeRange
    Dim targetWidth As Double, targetHeight As Double
    Dim doc As Document
    Dim pg As Page
    Dim layer As layer
    Dim shapeIndex As Long
    Dim totalPages As Integer
    Dim currentPage As Integer
    
    ' Disable events and save settings
    Optimization = True
    EventsEnabled = False
    ActiveDocument.SaveSettings
    ActiveDocument.PreserveSelection = False
    
    ' Initialize document and page count
    Set doc = ActiveDocument
    totalPages = doc.Pages.Count
    currentPage = 0
    
    ' Check if any shape is selected
    If activeSelection.Shapes.Count = 0 Then
        MsgBox "Oops! Nothing’s selected. Please select a powerclip frame and try again", vbExclamation, "No Selection"
        GoTo Cleanup
    End If
    
    ' Get the selected shape and its layer
    Set s = activeSelection.Shapes(1)
    Set layer = s.layer
    
    ' Find the index of the selected shape within its layer
    For shapeIndex = 1 To layer.Shapes.Count
        If layer.Shapes(shapeIndex) Is s Then
            Exit For
        End If
    Next shapeIndex
    
    ' Configure and display the status label
    gcSetStatus "Processing Page 0 of " & totalPages & " with index " & shapeIndex & "..."
    DoEvents ' Ensure initial UI update
    
    ' User confirmation
    ans = MsgBox("Going to proportionally fit content to frames with index " & shapeIndex & " on all pages and align to top," & vbNewLine & _
        "Do you want to continue?", vbQuestion + vbOKCancel, "Please confirm...")
    If ans = vbCancel Then
        gcIdle
        GoTo Cleanup
    End If
    
    ' Group commands for undo/redo efficiency
    doc.BeginCommandGroup "Image Adjustment"
    
    On Error GoTo Cleanup ' Ensure cleanup occurs even if an error happens
    
    ' Loop through all pages
    For Each pg In doc.Pages
        currentPage = currentPage + 1
        pg.Activate
        
        ' Update status with current page progress
        gcSetStatus "Processing Page " & currentPage & " of " & totalPages & " with index " & shapeIndex & "..."
        DoEvents ' Refresh UI
        
        ' Check if the layer exists on this page and has enough shapes
        If Not pg.Layers(layer.Name) Is Nothing Then
            Set layer = pg.Layers(layer.Name)
            If shapeIndex <= layer.Shapes.Count Then
                ' Get the shape at the same index on this page
                Set s = layer.Shapes(shapeIndex)
                
                ' Process the shape
                Set srPowerClipped = s.PowerClip.Shapes.FindShapes()
                
                targetWidth = s.SizeWidth
                targetHeight = s.SizeHeight
                
                ' Check if the frame has content and if content size is less than frame size
                If srPowerClipped.Count > 0 And (srPowerClipped.SizeWidth < targetWidth Or srPowerClipped.SizeHeight < targetHeight) Then
                    ' Calculate the scaling factor to fill the content
                    Dim scaleFactor As Double
                    If srPowerClipped.SizeWidth / srPowerClipped.SizeHeight > targetWidth / targetHeight Then
                        scaleFactor = targetHeight / srPowerClipped.SizeHeight
                    Else
                        scaleFactor = targetWidth / srPowerClipped.SizeWidth
                    End If
                    
                    ' Apply the scaling factor to fill the content
                    srPowerClipped.SetSize srPowerClipped.SizeWidth * scaleFactor, srPowerClipped.SizeHeight * scaleFactor
                    
                    ' Get the position and size of the shape (PowerClip frame)
                    Dim frameX As Double, frameY As Double, frameW As Double, frameH As Double
                    s.GetPosition frameX, frameY
                    s.GetSize frameW, frameH
                    
                    ' Calculate the new position for top alignment
                    Dim newX As Double
                    newX = frameX + (frameW - srPowerClipped.SizeWidth) / 2 ' Horizontally centered
                    Dim newY As Double
                    newY = frameY ' Align to the top of the frame
                    
                    ' Set the new position for the PowerClip content
                    srPowerClipped.SetPosition newX, newY
                End If
            End If
        End If
    Next pg
    
Cleanup:
    ' Finalize command group and restore settings
    doc.EndCommandGroup
    gcIdle
    ActiveDocument.PreserveSelection = True
    ActiveDocument.RestoreSettings
    EventsEnabled = True
    Optimization = False
    ActiveWindow.Refresh
    Application.Refresh
End Sub
Private Sub FitTextToFrame_Click()
    ' Disable unnecessary events and optimize performance
    Optimization = True
    EventsEnabled = False
    ActiveDocument.SaveSettings
    ActiveDocument.PreserveSelection = False

    Dim doc As Document
    Dim pg As Page
    Dim s As Shape
    Dim fontSize As Double
    Dim maxFontSize As Double
    Dim totalPages As Integer
    Dim currentPage As Integer
    
    ' Initialize document and page count
    Set doc = ActiveDocument
    totalPages = doc.Pages.Count
    currentPage = 0
    
    ' Configure and display the status label
    gcSetStatus "Processing Page 0 of " & totalPages
    DoEvents ' Ensure initial UI update
    
    ' Group commands for undo/redo efficiency
    doc.BeginCommandGroup "Text Adjustment"
    
    ' Iterate through pages
    For Each pg In doc.Pages
        currentPage = currentPage + 1
        pg.Activate
        
        ' Update status with current page progress
        gcSetStatus "Processing Page " & currentPage & " of " & totalPages
        DoEvents ' Refresh UI
        
        ' Process text shapes on the active layer
        For Each s In pg.ActiveLayer.Shapes
            If s.Type = cdrTextShape Then
                If s.text.Type = cdrParagraphText Then
                    ' Set initial font size
                    maxFontSize = s.text.FontProperties.Size
                    fontSize = maxFontSize
                    s.text.FontProperties.Size = fontSize
                    
                    ' Adjust font size if text overflows
                    While s.text.Overflow
                        fontSize = fontSize - 0.5
                        s.text.FontProperties.Size = fontSize
                    Wend
                End If
            End If
        Next s
    Next pg
    
    ' Finalize command group
    doc.EndCommandGroup
    
    ' Clean up and reset environment
    gcIdle
    ActiveDocument.PreserveSelection = True
    ActiveDocument.RestoreSettings
    EventsEnabled = True
    Optimization = False
    
    ' Refresh the UI
    ActiveWindow.Refresh
    Application.Refresh
End Sub

Private Sub folderpath_Change()

End Sub

Private Sub Frame1_Click()

End Sub

Private Sub Frame5_Click()

End Sub

Private Sub GetBitmapNoPC_Click()
 Dim selectedShape As Shape
    Dim layer As layer
    Dim shapeOrder As Long

    ' Check if any shape is selected
    If activeSelection.Shapes.Count > 0 Then
        ' Get the first selected shape
        Set selectedShape = activeSelection.Shapes(1)

        ' a QR / barcode made by Generator 1 or 2 is stored as -1 / -2
        If Left$(selectedShape.Name, 13) = "GeneratedCode" Then
            txtBitmapNumberPC.value = -Val(Mid$(selectedShape.Name, 14))
            Exit Sub
        End If

        ' Get the active layer
        Set layer = selectedShape.layer

        ' Find the order of the selected shape within the layer
        For bitmapOrder = 1 To layer.Shapes.Count
            If layer.Shapes(bitmapOrder) Is selectedShape Then
                ' Display the shape Number in a message box
                txtBitmapNumberPC.value = bitmapOrder
                Exit Sub
            End If
        Next bitmapOrder
    Else
        MsgBox "Oops! Nothing’s selected. Please select a image and try again.", vbExclamation, "No Selection"
    End If
End Sub

Private Sub GetPosionIdMaker3_Click()
Dim sh As Shape
    
    Dim sr As ShapeRange, w#, h#
    ActiveDocument.Unit = cdrInch
    Set sr = ActiveSelectionRange

 If sr.Count = 0 Then 'test that something is selected
      MsgBox "Plz select Barcode from your card template to get position and size", vbCritical
      Exit Sub
      
      Else
      bcX.value = activeSelection.centerX
      bcY.value = activeSelection.centerY
      sr(1).GetSize w, h
      bwth.value = w
      bhth.value = h
   End If
End Sub

Private Sub GetPosionIdMaker4_Click()
Dim sh As Shape
    
    Dim sr As ShapeRange, w#, h#
    ActiveDocument.Unit = cdrInch
    Set sr = ActiveSelectionRange

 If sr.Count = 0 Then 'test that something is selected
      MsgBox "Plz select image from your card template to get position and size", vbCritical
      Exit Sub
      
      Else
      
      ipX2.value = activeSelection.centerX
      ipY2.value = activeSelection.centerY
      sr(1).GetSize w, h
      iwth2.value = w
      ihth2.value = h
   End If
End Sub


Private Sub GetPositnIDMaker1_Click()
Dim sh As Shape
    
    Dim sr As ShapeRange, w#, h#
    ActiveDocument.Unit = cdrInch
    Set sr = ActiveSelectionRange

 If sr.Count = 0 Then 'test that something is selected
      MsgBox "Plz select image from your card template to get position and size", vbCritical
      Exit Sub
      
      Else
      
      ipX.value = activeSelection.centerX
      ipY.value = activeSelection.centerY
      sr(1).GetSize w, h
      iwth.value = w
      ihth.value = h
   End If
End Sub

Private Sub GetPositnIDMaker2_Click()
Dim sh As Shape
    
    Dim sr As ShapeRange, w#, h#
    ActiveDocument.Unit = cdrInch
    Set sr = ActiveSelectionRange

 If sr.Count = 0 Then 'test that something is selected
      MsgBox "Plz select image from your card template to get position and size", vbCritical
      Exit Sub
      
      Else
      
      ipX1.value = activeSelection.centerX
      ipY1.value = activeSelection.centerY
      sr(1).GetSize w, h
      iwth1.value = w
      ihth1.value = h
   End If
End Sub


Private Sub GetShapeNoForResize_Click()
Dim selectedShape As Shape
    Dim layer As layer
    Dim shapeOrder As Long

    ' Check if any shape is selected
    If activeSelection.Shapes.Count > 0 Then
        ' Get the first selected shape
        Set selectedShape = activeSelection.Shapes(1)

        ' Get the active layer
        Set layer = selectedShape.layer

        ' Find the order of the selected shape within the layer
        For shapeOrderT = 1 To layer.Shapes.Count
            If layer.Shapes(shapeOrderT) Is selectedShape Then
                ' Display the shape Number in a message box
                TargShape.value = shapeOrderT
                Exit Sub
            End If
        Next shapeOrderT
    Else
        MsgBox "Oops! Nothing’s selected. Please select a object and try again.", vbExclamation, "No Selection"
    End If
End Sub

Private Sub GetShapeNoPC_Click()

    Dim selectedShape As Shape
    Dim layer As layer
    Dim shapeOrder As Long

    ' Check if any shape is selected
    If activeSelection.Shapes.Count > 0 Then
        ' Get the first selected shape
        Set selectedShape = activeSelection.Shapes(1)

        ' Get the active layer
        Set layer = selectedShape.layer

        ' Find the order of the selected shape within the layer
        For shapeOrder = 1 To layer.Shapes.Count
            If layer.Shapes(shapeOrder) Is selectedShape Then
                ' Display the shape Number in a message box
                txtShapeNumberPC.value = shapeOrder
                Exit Sub
            End If
        Next shapeOrder
    Else
        MsgBox "Oops! Nothing’s selected. Please select a shape as frame and try again.", vbExclamation, "No Selection"
    End If
End Sub

Private Sub GetShapetoFrntNdBack_Click()
Dim selectedShape As Shape
    Dim layer As layer
    Dim shapeOrder As Long

    ' Check if any shape is selected
    If activeSelection.Shapes.Count > 0 Then
        ' Get the first selected shape
        Set selectedShape = activeSelection.Shapes(1)

        ' Get the active layer
        Set layer = selectedShape.layer

        ' Find the order of the selected shape within the layer
        For shapeOrderMFB = 1 To layer.Shapes.Count
            If layer.Shapes(shapeOrderMFB) Is selectedShape Then
                ' Display the shape Number in a message box
                MoveShapeBack.value = shapeOrderMFB
                Exit Sub
            End If
        Next shapeOrderMFB
    Else
        MsgBox "Oops! Nothing’s selected. Please select a object and try again.", vbExclamation, "No Selection"
    End If
End Sub

Private Sub Grouping_Click()
Optimization = True
EventsEnabled = False
ActiveDocument.SaveSettings
ActiveDocument.PreserveSelection = False
    ans = MsgBox("Going to Group everything in each Page of this Document," & vbNewLine & _
    "Do you want to Continue?", vbQuestion + vbOKCancel, "Please confirm...")
        If ans = vbCancel Then
            ActiveDocument.RestoreSettings
            EventsEnabled = True
            Optimization = False
            ActiveDocument.EndCommandGroup
            ActiveWindow.Refresh
            Application.Refresh
            Exit Sub
        End If
    ' Iterate through each page in the document
    For Each pg In ActiveDocument.Pages
        ' Activate the current page
        pg.Activate

        ' group all objects on the page
        ActivePage.Shapes.All.group
Next
MsgBox "Grouped Successfully", vbInformation, "Grouping"

ActiveDocument.PreserveSelection = True
ActiveDocument.RestoreSettings
EventsEnabled = True
Optimization = False
ActiveDocument.EndCommandGroup
ActiveWindow.Refresh
Application.Refresh

End Sub

Private Sub HeightR_Change()

End Sub

Private Sub ImgMove_Click()

 Dim iCtr As Long
 

    For iCtr = 0 To ListBox1.ListCount - 1
        If ListBox1.Selected(iCtr) = True Then
           ImgList.text = ListBox1.List(iCtr)
           ListBox1.RemoveItem (iCtr)
        End If
        Next iCtr
        
        ImgMove.Enabled = False
End Sub


Private Sub ImgMove1_Click()
Dim iCtr As Long
 

    For iCtr = 0 To ListBox1.ListCount - 1
        If ListBox1.Selected(iCtr) = True Then
           ImgList1.text = ListBox1.List(iCtr)
           ListBox1.RemoveItem (iCtr)
        End If
        Next iCtr
        
        ImgMove1.Enabled = False
End Sub


Private Sub ImgMove2_Click()
Dim iCtr As Long
 

    For iCtr = 0 To ListBox1.ListCount - 1
        If ListBox1.Selected(iCtr) = True Then
           ImgList2.text = ListBox1.List(iCtr)
           ListBox1.RemoveItem (iCtr)
        End If
        Next iCtr
        
        ImgMove2.Enabled = False
End Sub


Private Sub Label36_Click()

End Sub

Private Sub ListBox1_Click()

End Sub

Private Sub MoveToBack_Click()

End Sub

Private Sub PasteInEachPage_Click()
Optimization = True
EventsEnabled = False
ActiveDocument.SaveSettings
ActiveDocument.PreserveSelection = False

ans = MsgBox("Going to Paste Copied Data in to each Page of this Document," & vbNewLine & _
    "Do you want to Continue?", vbQuestion + vbOKCancel, "Please confirm...")
        If ans = vbCancel Then
            ActiveDocument.RestoreSettings
            EventsEnabled = True
            Optimization = False
            ActiveDocument.EndCommandGroup
            ActiveWindow.Refresh
            Application.Refresh
            Exit Sub
        End If
    
Dim pasteopt As StructPasteOptions
    Dim Paste1 As ShapeRange
    
    For i = 1 To ActiveDocument.Pages.Count
        ActiveDocument.Pages(i).Activate
        Set pasteopt = CreateStructPasteOptions
        Set Paste1 = ActiveLayer.PasteEx(pasteopt)
    Next i
    
    MsgBox "Pasted Successfully", vbInformation, "Pasted"
    
    ActiveDocument.PreserveSelection = True
ActiveDocument.RestoreSettings
EventsEnabled = True
Optimization = False
ActiveDocument.EndCommandGroup
ActiveWindow.Refresh
Application.Refresh

End Sub

Private Sub ResetCardMaker_Click()
ListBox1.Clear
Set gcFields = Nothing
ImgList.value = ""
ImgList1.value = ""
ImgList2.value = ""
BarList.value = ""
ipX.value = "0"
ipY.value = "0"
iwth.value = "0"
ihth.value = "0"
ipX1.value = "0"
ipY1.value = "0"
iwth1.value = "0"
ihth1.value = "0"
ipX2.value = "0"
ipY2.value = "0"
iwth2.value = "0"
ihth2.value = "0"
bcX.value = "0"
bcY.value = "0"
bwth.value = "0"
bhth.value = "0"
Me.imgCheck.value = False
Me.imgCheck1.value = False
Me.imgCheck2.value = False
Me.bcCheck.value = False
Me.folderpath.value = "Choose a TXT file"


ImgMove.Enabled = True
Me.ImgMove1.Enabled = True
Me.ImgMove2.Enabled = True
Me.BarcodeMove.Enabled = True
End Sub
Sub DupPage()
    Dim sr As ShapeRange
    Dim pCur As Page
    Dim pNext As Page
    Dim sDuplicate As Shape
    Dim lyr As Layer
    Dim dupRange As ShapeRange

    Set pCur = ActivePage

    If GC_FAST_PAGES Then
        ' experimental fast mode: copy a whole layer at once instead of shape by shape
        Set pNext = ActiveDocument.InsertPages(1, False, pCur.index)
        For Each lyr In pCur.Layers
            If lyr.Shapes.Count > 0 Then
                Set dupRange = lyr.Shapes.All.Duplicate
                dupRange.ReverseRange.MoveToLayer pNext.Layers(lyr.Name)
            End If
        Next lyr
        pNext.Activate
        Exit Sub
    End If

    Set sr = ActivePage.Shapes.All
    Set pNext = ActiveDocument.InsertPages(1, False, ActivePage.index)

    For Each s In sr.ReverseRange
        Set sDuplicate = s.Duplicate
        sDuplicate.MoveToLayer pNext.Layers(s.layer.Name)
    Next s

    pNext.Activate
End Sub
Function FileExists(filePath As String) As Boolean
Dim TestStr As String
    TestStr = ""
    On Error Resume Next
    TestStr = Dir(filePath)
    On Error GoTo 0
    If TestStr = "" Then
        FileExists = False
    Else
        FileExists = True
    End If
End Function
Sub AlignAndPowerClip()
    On Error Resume Next
    ActiveDocument.Unit = cdrCentimeter

   

    ' Validate the input
    Dim shapeNumber As Integer
    Dim bitmapNumber As Integer
    shapeNumber = txtShapeNumberPC.value
    bitmapNumber = txtBitmapNumberPC.value

    If shapeNumber <= 0 Or bitmapNumber = 0 Then
        MsgBox "Invalid shape or bitmap number. Exiting macro.", vbExclamation
        Exit Sub
    End If

    ' Iterate through each page in the document
    For Each pg In ActiveDocument.Pages
        ' Activate the current page
        pg.Activate

        ' Set the specified shape for alignment and power-clipping
        Dim shapeToAlignAndPowerClip As Shape
        Set shapeToAlignAndPowerClip = ActivePage.ActiveLayer.Shapes(shapeNumber)

        ' Get the position and size of the specified shape
        Dim x As Double, y As Double, w As Double, h As Double
        shapeToAlignAndPowerClip.GetPosition x, y
        shapeToAlignAndPowerClip.GetSize w, h

        ' Iterate through each shape in the active layer
        Dim bitmapCount As Integer
        bitmapCount = 0

        If bitmapNumber < 0 Then
            ' QR / barcode made by Generator 1 or 2: top centre of the frame, size unchanged
            Dim qrShape As Shape
            Set qrShape = gcFindCode(-bitmapNumber)
            If Not qrShape Is Nothing Then
                qrShape.CenterX = shapeToAlignAndPowerClip.CenterX
                qrShape.TopY = shapeToAlignAndPowerClip.TopY
                qrShape.AddToPowerClip shapeToAlignAndPowerClip, cdrFalse
            End If
        Else
        For Each s In ActivePage.ActiveLayer.Shapes
            ' Check if the shape is a bitmap
            If s.Type = cdrBitmapShape Then
                bitmapCount = bitmapCount + 1

                ' Check if the current bitmap is the one to align and power-clip
                If bitmapCount = bitmapNumber Then
                    ' Calculate the new position for top center alignment
                    Dim newX As Double
                    newX = x + (w - s.SizeWidth) / 2

                    ' Set the new position for the bitmap
                    s.SetPosition newX, y
                    s.SetSize s.SizeWidth, s.SizeHeight ' Optional: Resize the bitmap if needed

                    ' Get the position of the bitmap before power-clipping
                    Dim bmpX As Double, bmpY As Double
                    s.GetPosition bmpX, bmpY

                    ' Power-clip the bitmap to the specified shape
                    s.SetBoundingBox x, y, w, False
                    s.AddToPowerClip shapeToAlignAndPowerClip, cdrTrue

                    ' Set the position of the power-clipped bitmap
                    Dim powerClipBmp As Shape
                    Set powerClipBmp = shapeToAlignAndPowerClip.PowerClip.Shapes(1)
                    powerClipBmp.SetPosition bmpX, bmpY

                    Exit For
                End If
            End If
        Next s
        End If
    Next pg

End Sub

Sub PowerClipOnly()
    On Error Resume Next
    ActiveDocument.Unit = cdrCentimeter


    ' Validate the input
    Dim shapeNumber As Integer
    Dim bitmapNumber As Integer
    shapeNumber = txtShapeNumberPC.value
    bitmapNumber = txtBitmapNumberPC.value

    If shapeNumber <= 0 Or bitmapNumber = 0 Then
        MsgBox "Invalid shape or bitmap number. Exiting macro.", vbExclamation
        Exit Sub
    End If


    ' Iterate through each page in the document
    For Each pg In ActiveDocument.Pages
        ' Activate the current page
        pg.Activate

        ' Set the specified shape for power-clipping
        Dim shapeToPowerClip As Shape
        Set shapeToPowerClip = ActivePage.ActiveLayer.Shapes(shapeNumber)

        ' Get the position and size of the specified shape
        Dim x As Double, y As Double, w As Double, h As Double
        shapeToPowerClip.GetPosition x, y
        shapeToPowerClip.GetSize w, h

        ' Iterate through each shape in the active layer
        Dim bitmapCount As Integer
        bitmapCount = 0

        If bitmapNumber < 0 Then
            ' QR / barcode made by Generator 1 or 2: centre of the frame, size unchanged
            Dim qrShape As Shape
            Set qrShape = gcFindCode(-bitmapNumber)
            If Not qrShape Is Nothing Then
                qrShape.CenterX = shapeToPowerClip.CenterX
                qrShape.CenterY = shapeToPowerClip.CenterY
                qrShape.AddToPowerClip shapeToPowerClip, cdrFalse
            End If
        Else
        For Each s In ActivePage.ActiveLayer.Shapes
            ' Check if the shape is a bitmap
            If s.Type = cdrBitmapShape Then
                bitmapCount = bitmapCount + 1

                ' Check if the current bitmap is the one to power-clip
                If bitmapCount = bitmapNumber Then
                    ' Power-clip the bitmap to the specified shape
                    s.SetBoundingBox x, y, w, False
                    s.AddToPowerClip shapeToPowerClip, cdrTrue
                    Exit For
                End If
            End If
        Next s
        End If
    Next pg
End Sub

Private Sub RotateObj_Click()
    Dim doc As Document
    Dim pg As Page
    Dim layer As layer
    Dim selectedShape As Shape
    Dim shapeOrder As Long
    Dim totalPages As Integer
    Dim currentPage As Integer
    
    ' Disable events and save settings
    Optimization = True
    EventsEnabled = False
    ActiveDocument.SaveSettings
    ActiveDocument.PreserveSelection = False
    
    ' Initialize document and page count
    Set doc = ActiveDocument
    totalPages = doc.Pages.Count
    currentPage = 0
    
    ' Check if any shape is selected
    If activeSelection.Shapes.Count = 0 Then
        MsgBox "Oops! Nothing’s selected. Please select an object and try again.", vbExclamation, "No Selection"
        GoTo Cleanup
    End If
    
    ' Get the first selected shape and its layer
    Set selectedShape = activeSelection.Shapes(1)
    Set layer = selectedShape.layer
    
    ' Find the order of the selected shape within the layer
    For shapeOrder = 1 To layer.Shapes.Count
        If layer.Shapes(shapeOrder) Is selectedShape Then
            Exit For
        End If
    Next shapeOrder
    
    ' Configure and display the status label
    gcSetStatus "Processing Page 0 of " & totalPages & " with shape order " & shapeOrder & "..."
    DoEvents ' Ensure initial UI update
    
    ' Ask for confirmation
    ans = MsgBox("Going to Rotate 180 Selected Object in each Page of this Document," & vbNewLine & _
        "Do you want to Continue?", vbQuestion + vbOKCancel, "Please confirm...")
    If ans = vbCancel Then
        gcIdle
        GoTo Cleanup
    End If
    
    ' Group commands for undo/redo efficiency
    doc.BeginCommandGroup "Rotate Objects"
    
    On Error GoTo Cleanup ' Ensure cleanup occurs even if an error happens
    
    ' Loop through all pages
    For Each pg In doc.Pages
        currentPage = currentPage + 1
        pg.Activate
        
        ' Update status with current page progress
        gcSetStatus "Processing Page " & currentPage & " of " & totalPages & " with shape order " & shapeOrder & "..."
        DoEvents ' Refresh UI
        
        ' Check if the layer exists on this page and has enough shapes
        If Not pg.Layers(layer.Name) Is Nothing Then
            Set layer = pg.Layers(layer.Name)
            If shapeOrder <= layer.Shapes.Count Then
                ' Rotate the shape at the same order on this page
                layer.Shapes(shapeOrder).Rotate 180
            End If
        End If
    Next pg
    
Cleanup:
    ' Finalize command group and restore settings
    doc.EndCommandGroup
    gcIdle
    ActiveDocument.PreserveSelection = True
    ActiveDocument.RestoreSettings
    EventsEnabled = True
    Optimization = False
    ActiveWindow.Refresh
    Application.Refresh
End Sub
Private Sub RotateShapeByOrder(ByVal shapeOrder As Long)
    Dim currentPage As Page
    For Each currentPage In ActiveDocument.Pages
        If shapeOrder <= currentPage.Shapes.Count Then
            Dim selectedShape As Shape
            Set selectedShape = currentPage.Shapes(shapeOrder)

            ' Get the original position of the shape
            Dim originalX As Double
            Dim originalY As Double
            originalX = selectedShape.PositionX
            originalY = selectedShape.PositionY

            ' Get the center coordinates of the shape
            Dim centerX As Double
            Dim centerY As Double
            centerX = selectedShape.centerX
            centerY = selectedShape.centerY

            ' Set the rotation point to the center of the shape
            selectedShape.RotationCenterX = centerX
            selectedShape.RotationCenterY = centerY

            ' Rotate the shape
            selectedShape.RotateEx 180#, 3.23356, 14.60857

            ' Reset the position of the shape
            selectedShape.PositionX = originalX
            selectedShape.PositionY = originalY
        End If
    Next currentPage
End Sub

Private Sub Status_Click()

End Sub

Private Sub TargShape_Change()

End Sub

Private Sub UserForm_Click()

End Sub



Private Sub GrpGetBoundingBox_Click()
    Dim sr As ShapeRange, allSr As ShapeRange
    Dim i As Long

    If Documents.Count = 0 Then Exit Sub
    Set sr = ActiveSelectionRange
    If sr.Count = 0 Then
        MsgBox "Please select the shapes you want to group, then click Capture.", vbExclamation
        Exit Sub
    End If

    Set SourcePage = ActivePage
    Set SelectedShapeIndices = New Collection
    Set allSr = SourcePage.Shapes.All
    For i = 1 To allSr.Count
        If allSr.Item(i).Selected Then SelectedShapeIndices.Add i
    Next i

    If SelectedShapeIndices.Count = 0 Then
        MsgBox "These shapes could not be captured. Select whole objects (not objects inside a group) and try again.", vbExclamation
        Exit Sub
    End If

    Me.GrpMinX.Value = SelectedShapeIndices(1)                        ' first index
    Me.GrpMaxX.Value = SelectedShapeIndices(SelectedShapeIndices.Count) ' last index
    Me.GrpMinY.Value = SelectedShapeIndices.Count                     ' number of shapes
    Me.GrpMaxY.Value = SourcePage.index                               ' page number
    gcNote SelectedShapeIndices.Count & " shapes captured from page " & SourcePage.index & "."
End Sub

Private Sub GrpCmdGroupShapes_Click()
    Dim doc As Document, pg As Page, allSr As ShapeRange, sr As ShapeRange, shp As Shape
    Dim i As Long, idx As Long, pgNo As Long, totalPages As Long, nGrouped As Long
    Dim failPages As String, hadError As Boolean

    If SelectedShapeIndices Is Nothing Then
        MsgBox "Nothing captured yet. Select the shapes on one page and click Capture first.", vbExclamation
        Exit Sub
    End If
    If SelectedShapeIndices.Count < 2 Then
        MsgBox "Capture at least 2 shapes first (select them all, then click Capture).", vbExclamation
        Exit Sub
    End If

    Set doc = ActiveDocument
    totalPages = doc.Pages.Count
    hadError = False

    Optimization = True
    EventsEnabled = False
    doc.PreserveSelection = False
    doc.BeginCommandGroup "Group Shapes by Indices"
    On Error GoTo EH

    For Each pg In doc.Pages
        pgNo = pgNo + 1
        pg.Activate
        Set sr = New ShapeRange
        Set allSr = pg.Shapes.All
        For i = 1 To SelectedShapeIndices.Count
            idx = SelectedShapeIndices(i)
            If idx <= allSr.Count Then
                Set shp = allSr.Item(idx)
                If shp.Visible And Not shp.Locked Then sr.Add shp
            End If
        Next i

        If sr.Count > 1 Then
            On Error Resume Next
            sr.Group
            If Err.Number = 0 Then
                nGrouped = nGrouped + 1
            Else
                If Len(failPages) < 150 Then failPages = failPages & pgNo & ", "
                Err.Clear
            End If
            On Error GoTo EH
        Else
            If Len(failPages) < 150 Then failPages = failPages & pgNo & ", "
        End If

        If (pgNo Mod 5) = 0 Or pgNo = totalPages Then
            gcSetStatus "Grouping page " & pgNo & " of " & totalPages
            DoEvents
        End If
    Next pg

Finish:
    gcIdle
    doc.PreserveSelection = True
    EventsEnabled = True
    Optimization = False
    doc.EndCommandGroup
    ActiveWindow.Refresh
    Application.Refresh
    If hadError Then Exit Sub

    If failPages = "" Then
        MsgBox "Shapes grouped on all " & totalPages & " pages.", vbInformation, "Operation Complete"
    Else
        MsgBox "Grouped on " & nGrouped & " of " & totalPages & " pages." & vbCrLf & _
               "Not grouped (shapes not found, hidden or locked) on page(s): " & Left$(failPages, Len(failPages) - 2), _
               vbExclamation, "Operation Complete"
    End If
    Exit Sub
EH:
    hadError = True
    MsgBox "An error occurred: " & Err.Description, vbCritical
    Resume Finish
End Sub

' Helper function to get shape type name
Private Function GetShapeTypeName(shapeType As Long) As String
    Select Case shapeType
        Case 1: GetShapeTypeName = "Rectangle"
        Case 2: GetShapeTypeName = "Ellipse"
        Case 6: GetShapeTypeName = "Text"
        Case 5: GetShapeTypeName = "Curve"
        Case 7: GetShapeTypeName = "Group"
        Case Else: GetShapeTypeName = "Other (" & shapeType & ")"
    End Select
End Function

Private Sub UserForm_Initialize()
    ' Initialize variables
    Set SelectedShapeIndices = New Collection
    Set SourcePage = Nothing
    ' Initialize text boxes
    Me.GrpMinX.value = ""
    Me.GrpMaxX.value = ""
    Me.GrpMinY.value = ""
    Me.GrpMaxY.value = ""
    ' Initialize status label
    gcSetStatus "Ready"
    gcBuildUi
End Sub
