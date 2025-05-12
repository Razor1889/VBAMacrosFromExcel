Attribute VB_Name = "Module1"
' Declare the WaitMessage function for delay functionality
Private Declare PtrSafe Function WaitMessage Lib "user32" () As Long

' Define the Wait function (for delay in seconds)
Public Sub Wait(Seconds As Double)
    Dim endtime As Double
    endtime = Timer + Seconds
    Do
        WaitMessage
        DoEvents
    Loop While Timer < endtime
End Sub

' Main subroutine to move the boat and display the GIFs
Sub MoveBoatFromCSV()
    Dim ws As Slide
    Dim csvPath As String
    Dim fileNumber As Integer
    Dim lineData As String
    Dim values() As String
    Dim x As Integer, y As Integer
    Dim pwm As Integer, turnAngle As Integer, currAngle As Integer
    Dim shape As shape
    Dim cellName As String
    Dim boatImg As shape
    Dim infoText As shape
    Dim gifShape As shape
    Dim fso As Object
    
    ' Set the slide (Active slide in PowerPoint)
    Set ws = ActivePresentation.Slides(1)
    
    ' Define file path (Ensure CSV is in the same folder as the PowerPoint)
    csvPath = ActivePresentation.Path & "\data.csv"
    
    ' Create FileSystemObject to check file existence
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Check if the file exists
    If Not fso.FileExists(csvPath) Then
        MsgBox "CSV file not found at: " & csvPath
        Exit Sub
    End If
    
    ' Open the CSV file
    On Error GoTo FileError
    fileNumber = FreeFile
    Open csvPath For Input As #fileNumber
    On Error GoTo 0
    
    ' Get boat and text shapes
    On Error Resume Next
    Set boatImg = ws.Shapes("BoatImage")
    Set infoText = ws.Shapes("InfoText")
    On Error GoTo 0
    
    If boatImg Is Nothing Or infoText Is Nothing Then
        MsgBox "Error: Missing BoatImage or InfoText. Please check shape names.", vbCritical
        Exit Sub
    End If
    
    ' Read the CSV line by line
    Do While Not EOF(fileNumber)
        Line Input #fileNumber, lineData
        values = Split(lineData, ",")
        
        ' Parse values
        x = CInt(values(0))
        y = CInt(values(1))
        pwm = CInt(values(2))
        turnAngle = CInt(values(3))
        currAngle = CInt(values(4))
        
        ' Find the corresponding cell
        cellName = "Cell_" & x & "_" & y
        Set shape = Nothing
        
        On Error Resume Next
        Set shape = ws.Shapes(cellName)
        On Error GoTo 0
        
        ' Ensure the shape exists
        If Not shape Is Nothing Then
            ' Move the boat image to this cell
            boatImg.Top = shape.Top
            boatImg.Left = shape.Left
        End If
        
        ' Update PWM, angles in InfoText
        infoText.TextFrame.TextRange.Text = "PWM: " & pwm & vbCrLf & _
                                            "Turn Angle: " & turnAngle & vbCrLf & _
                                            "Current Angle: " & currAngle
        
        ' Hide all GIFs at the start of each iteration
        HideAllGIFs ws
        
        ' Display appropriate GIF for turns
        Set gifShape = Nothing
        On Error Resume Next
        Select Case turnAngle
            Case 45
                Set gifShape = ws.Shapes("Turn_45")
            Case -45
                Set gifShape = ws.Shapes("Turn_-45")
            Case 180
                Set gifShape = ws.Shapes("Turn_180")
            Case 0  ' New case for straight movement
                Set gifShape = ws.Shapes("Straight_Move")
        End Select
        On Error GoTo 0
        
        If Not gifShape Is Nothing Then
            gifShape.Visible = msoTrue
        End If
        
        ' Short delay for animation effect (1 second)
        Call Wait(1) ' Delay in seconds (1 second)
    Loop
    
    ' Close the CSV file
    Close #fileNumber
    Exit Sub
    
FileError:
    MsgBox "Error opening file: " & csvPath, vbCritical
End Sub

' Function to hide all GIFs on the slide
Sub HideAllGIFs(ws As Slide)
    On Error Resume Next
    ws.Shapes("Turn_45").Visible = msoFalse
    ws.Shapes("Turn_-45").Visible = msoFalse
    ws.Shapes("Turn_180").Visible = msoFalse
    ws.Shapes("Straight_Move").Visible = msoFalse  ' Hide straight move GIF
    On Error GoTo 0
End Sub

' Reset subroutine to set all values to zero and make boat go back to Cell_1_1
Sub ResetAll()
    Dim ws As Slide
    Dim boatImg As shape
    Dim infoText As shape
    Dim gifShape As shape
    Dim shape As shape
    Dim cellName As String
    
    ' Set the slide (Active slide in PowerPoint)
    Set ws = ActivePresentation.Slides(1)
    
    ' Get boat and text shapes
    On Error Resume Next
    Set boatImg = ws.Shapes("BoatImage")
    Set infoText = ws.Shapes("InfoText")
    On Error GoTo 0
    
    If boatImg Is Nothing Or infoText Is Nothing Then
        MsgBox "Error: Missing BoatImage or InfoText. Please check shape names.", vbCritical
        Exit Sub
    End If
    
    ' Reset the boat position to Cell_1_1
    cellName = "Cell_1_1"
    Set shape = Nothing
    On Error Resume Next
    Set shape = ws.Shapes(cellName)
    On Error GoTo 0
    
    If Not shape Is Nothing Then
        boatImg.Top = shape.Top
        boatImg.Left = shape.Left
    End If
    
    ' Reset infoText
    infoText.TextFrame.TextRange.Text = "PWM: 0" & vbCrLf & _
                                        "Turn Angle: 0 °" & vbCrLf & _
                                        "Current Angle: 0 °"
    
    ' Hide all GIFs
    HideAllGIFs ws
End Sub


