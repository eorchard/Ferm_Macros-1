'Makes string comparisons case insensitive
Option Compare Text
 
'Quickly generates a plate in our typical format
Sub Quick_Populate()
    Dim numberOfVessels As Integer
    Dim counter As Integer
    Dim firstRow As Integer
    Dim DG_Unit As String
    Dim wrapped As Boolean
    Dim emptyRightTwoColumns As Boolean
    Dim exitFor As Boolean
    counter = 0
    firstRow = 2 'First row of the table
 
    'Collect user input
    numberOfVessels = Range("Y9").Value
    DG_Unit = Range("Y10").Value
    numberOfTimepoints = Range("Y11").Value
    emptyRightTwoColumns = IIf(Range("Y26").Value = "Yes", True, False)
   
    Call Clear_Table
   
    'Populate Timepoints
    For x = 0 To numberOfVessels - 1
        For y = 0 To numberOfTimepoints - 1
            If Not Wrapped Then
               'Wrap plate to fourth row
                If (counter = 12) Or (counter = 10 And emptyRightTwoColumns) Then
                    firstRow = firstRow + 3
                    counter = 0
                    wrapped = True
                'Error notification for excessive samples
            Else
                ElseIf (counter = 12) Or (counter = 10 And emptyRightTwoColumns) Then
                    MsgBox "Plate Overflow! Please adjust Quick Populate settings."
                    'Break loop
                    exitFor = True
                    Exit For
                End If    
            End If

            For z = 0 to 2
                Range("C" & (firstRow + z + (8 * counter))).Value = DG_Unit & x + 1 'Vessel
                Range("D" & (firstRow + z + (8 * counter))).Value = "I" & Range("Y" & (12 + y)).Value 'Timepoint
                Range("E" & (firstRow + z + (8 * counter))).Value = Range("Y" & (18 + x)).Value 'Color
            Next
            counter = counter + 1
        Next
        'Break loop
        If exitFor Then Exit For
    Next
   
    'Wrapped state will determine whether CFB ends up on Row D or Row G
    firstRow = IIf(wrapped, 8, 5)
   
    'Populate CFB
    For x = 0 To numberOfVessels - 1
        For y = 0 to 1
            Range("C" & firstRow + y + (8 * x)).Value = DG_Unit & x + 1 'Vessel
            Range("D" & firstRow + y + (8 * x)).Value = "CFB" 'Timepoint
            Range("E" & firstRow + y + (8 * x)).Value = Range("Y" & (18 + x)).Value 'Color
        Next
    Next
   
    Call Generate_Plate
   
End Sub
 
Sub Generate_Plate()
    Dim ovalShape As Shape
    Dim A1_X As Double
    Dim A1_Y As Double
    Dim colorArray(96) As String
    Dim timepointArray(96) As String
    Dim dasgipArray(96) As String
    Dim color As String
    Dim counter As Integer
    Dim x As Integer
     
    'Clear plate
    Call Clear_Plate
     
    'Absolute location of well A1, rest will be calculated relatively
    Set plateLocation = ActiveSheet.Shapes("Plate Map")
    A1_X = 397.5
    A1_Y = 60.75
     
    'Initialize plate map in correct location with correct size
    With ActiveSheet.Shapes("Plate Map")
        .Left = 366
        .Top = 1.25
        .Height = 378.72
        .Width = 510.48
    End With
     
    'Populate arrays with values, begin at x+1 since the first row is a header
    For x = 1 To 96
        dasgipArray(x) = Cells(x + 1, 3).Value
        timepointArray(x) = Cells(x + 1, 4).Value
        colorArray(x) = Cells(x + 1, 5).Value
    Next
       
    'Populate wells
    counter = 1
     
    For i = 1 To 12
        For j = 1 To 8
            color = colorArray(counter)
           
            'Skip elements in colorArray that are empty
            If colorArray(counter) <> "" Then
                Set ovalShape = ActiveSheet.Shapes(color).Duplicate
                With ovalShape
                    .Name = color & "_Copy"
                    .Left = A1_X
                    .Top = A1_Y
                   
                    'Color wells relative to A1
                    .IncrementLeft 38.75 * (i - 1)
                    .IncrementTop 38.25 * (j - 1)
                   
                    'Populate wells with text
                    With .TextFrame
                        .MarginLeft = 0
                        .MarginRight = 0
                        .MarginTop = 0
                        .MarginBottom = 0
                        .Characters.Text = dasgipArray(counter) & vbNewLine & timepointArray(counter)
                        .Characters.Font.Size = 9.5
                        .Characters.Font.Bold = True
                        .HorizontalAlignment = xlHAlignCenter
                    End With
                End With
            End If
            counter = counter + 1
        Next
    Next
End Sub
 
'Clear pre-existing data
Sub Clear_Table()
    Range("C2:E200").ClearContents
End Sub
