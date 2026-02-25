Sub SetupCanvasMinimal()
    Dim frameWidth As Long, frameHeight As Long
    Dim pixelSize As Double
    
    ' Set to match your Python output
    frameWidth = 100
    frameHeight = 100
    
    Application.ScreenUpdating = False
    
    ' 1. Clear everything (formats and text)
    Cells.ClearFormats
    Cells.ClearContents
    
    ' 2. The Perfect Square Math
    With Range(Cells(1, 1), Cells(frameHeight, frameWidth))
        .ColumnWidth = 2
    End With
    
    pixelSize = Cells(1, 1).Width
    Range(Cells(1, 1), Cells(frameHeight, frameWidth)).RowHeight = pixelSize
    
    ' 3. Paint the entire canvas black instantly in one command
    Range(Cells(1, 1), Cells(frameHeight, frameWidth)).Interior.Color = RGB(0, 0, 0)
    
    Application.ScreenUpdating = True
    MsgBox "Blank canvas ready!", vbInformation
End Sub




Sub PlayVideoStreamMaxSpeed()
    Dim x As Long, y As Long
    Dim fileNum As Integer
    Dim frameWidth As Long, frameHeight As Long
    Dim binFilePath As Variant
    Dim frameData() As Byte
    Dim byteIndex As Long
    Dim ws As Worksheet
    
    Set ws = ActiveSheet
    frameWidth = 100
    frameHeight = 100
    
    ReDim frameData(0 To (frameWidth * frameHeight * 3) - 1)
    
    binFilePath = Application.GetOpenFilename("Binary Files (*.bin), *.bin")
    If binFilePath = False Then Exit Sub
    
    ' --- THE ULTIMATE SPEED HACKS ---
    ' Turn off all background Excel operations
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayStatusBar = False
    
    MsgBox "Press CTRL+BREAK to stop the video.", vbInformation
    
    Do While True
        fileNum = FreeFile
        Open binFilePath For Binary Access Read As #fileNum
        
        Do While Not EOF(fileNum)
            Application.ScreenUpdating = False
            
            Get #fileNum, , frameData
            byteIndex = 0
            
            ' Paint cells
            For y = 1 To frameHeight
                For x = 1 To frameWidth
                    ws.Cells(y, x).Interior.Color = RGB( _
                        frameData(byteIndex), _
                        frameData(byteIndex + 1), _
                        frameData(byteIndex + 2))
                        
                    byteIndex = byteIndex + 3
                Next x
            Next y
            
            Application.ScreenUpdating = True
            DoEvents
        Loop
        Close #fileNum
    Loop
End Sub
