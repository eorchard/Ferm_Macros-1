'Makes string comparisons case insensitive
Option Compare Text

'Quickly generates a plate in our typical format
Sub Quick_Populate()
    Dim numberOfVessels As Integer, counter As Integer, firstRow As Integer
    Dim DG_Unit As String
    Dim wrapped As Boolean, emptyRightTwoColumns As Boolean, exitFor As Boolean
    firstRow = 2 'First row of the table

    'Collect user input
    numberOfVessels = Range("Y9").Value
    DG_Unit = Range("Y10").Value
    numberOfTimepoints = Range("Y11").Value
    emptyRightTwoColumns = IIf(ActiveSheet.Shapes("isEmptyRightTwoChecked").ControlFormat.Value = 1, True, False)
  
    'Clear table
    Call Clear_Table
   
    'Loop over vessels and timepoints
    For x = 0 To numberOfVessels - 1
        For y = 0 To numberOfTimepoints - 1

            'Check if plate needs to be wrapped to fourth row
            If (counter = 12) Or (counter = 10 And emptyRightTwoColumns) Then
                If Not wrapped Then
                    firstRow = firstRow + 3
                    counter = 0
                    wrapped = True

                'Error notification for excessive samples
               Else
                    MsgBox "Plate Overflow! Please adjust Quick Populate settings."
                    'Break inner loop if overflowed
                    exitFor = True
                    Exit For
                End If
            End If

            'Populate timepoints
            For Z = 0 To 2
                'Vessel
                Range("C" & (firstRow + Z + (8 * counter))).Value = DG_Unit & x + 1
                'Timepoint
                Range("D" & (firstRow + Z + (8 * counter))).Value = "I" & Range("Y" & (12 + y)).Value
                'Color
                Range("E" & (firstRow + Z + (8 * counter))).Value = Range("Y" & (18 + x)).Value
            Next
            counter = counter + 1
        Next
        'Break outer loop if overflowed
        If exitFor Then Exit For
    Next
   
    'Wrapped state will determine whether CFB ends up on Row D or Row G
    firstRow = IIf(wrapped, 8, 5)
   
    'Populate CFB
    For x = 0 To numberOfVessels - 1
        For y = 0 To 1
            'Vessel
            Range("C" & firstRow + y + (8 * x)).Value = DG_Unit & x + 1
            'Timepoint
            Range("D" & firstRow + y + (8 * x)).Value = "CFB"
            'Color
            Range("E" & firstRow + y + (8 * x)).Value = Range("Y" & (18 + x)).Value
        Next
    Next
   
    Call Generate_Plate
   
End Sub

Sub Generate_Plate()
    Dim ovalShape As Shape
    Dim A1_X As Double, A1_Y As Double
    Dim color As String, colorArray(96) As String, timePointArray(96) As String, dasgipArray(96) As String
    Dim counter As Integer
     
    'Clear plate
    Call Clear_Plate
     
    'Initialize plate map in correct location with correct size
    With ActiveSheet.Shapes("Plate Map")
        .Left = Range("I1").Left
        .Top = Range("A3").Top
        .Height = Range("I1:I24").Height
    End With

    'Generate plate header
    Range("I1").Value = Range("Y36").Value & Space(16) & "FME" & Range("Y34").Value
         
    'Absolute location of well A1, rest will be calculated relatively
    Set plateLocation = ActiveSheet.Shapes("Plate Map")
    A1_X = plateLocation.Left + 40
    A1_Y = plateLocation.Top + 30
     
    'Populate arrays with values, begin at x+1 since the first row is a header
    For x = 1 To 96
        dasgipArray(x) = Cells(x + 1, 3).Value
        timePointArray(x) = Cells(x + 1, 4).Value
        colorArray(x) = Cells(x + 1, 5).Value
        
    'Generate CompoundID column to facilitate CBIS fraction plate import. Example: FME0376-DG3_u5-I6-001
        If timePointArray(x) <> "" Then
            If timePointArray(x) = "CFB" Then
            'Replace "CFB" string with the last timepoint
                Range("F" & (x + 1)).Value = "FME" & Range("Y34").Value & "-DG" & Mid(Cells(x + 1, 3).Value, 3, 1) & "_u" & Right(Cells(x + 1, 3).Value, 1) & "-I" & _
                    (Range("Y" & (11 + Range("Y11").Value)).Value) & "-002"
            Else
                Range("F" & (x + 1)).Value = "FME" & Range("Y34").Value & "-DG" & Mid(Cells(x + 1, 3).Value, 3, 1) & "_u" & Right(Cells(x + 1, 3).Value, 1) & "-" & _
                    Cells(x + 1, 4).Value & "-001"
            End If
        End If
    Next
       
    'Populate wells
    counter = 1

    'Loop over 96-well plate
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
                        .Characters.Text = dasgipArray(counter) & vbNewLine & timePointArray(counter)
                        .Characters.Font.Size = Range("Y35").Value
                        .Characters.Font.Bold = True
                        .HorizontalAlignment = xlHAlignCenter
                    End With
                End With
            Else
                'Create empty wells
                Set ovalShape = ActiveSheet.Shapes("Empty").Duplicate
                With ovalShape
                    .Name = "Empty_Copy"
                    .Left = A1_X
                    .Top = A1_Y
                    .IncrementLeft 38.75 * (i - 1)
                    .IncrementTop 38.25 * (j - 1)
                End With
            End If
            counter = counter + 1
        Next
    Next
End Sub

Sub Custom_Populate()
    Dim customTableCounter As Integer, customPopulateFirstRow As Integer, customPopulateLastRow As Integer, customPopulateFirstColumn As Integer, customPopulateLastColumn As Integer, _
        counter As Integer
    Dim customTimePointArray(96) As String, customDasgipArray(96) As String, customColorArray(96) As String
    Dim emptyRightTwoColumns As Boolean
    
    customPopulateFirstRow = 41
    customPopulateLastRow = 64
    customPopulateFirstColumn = 23
    customPopulateLastColumn = 27
    customTableCounter = 1
    counter = 0
    firstRow = 1
    emptyRightTwoColumns = IIf(ActiveSheet.Shapes("isEmptyRightTwoCheckedCustom").ControlFormat.Value = 1, True, False)
    
    'Collect user input
    For x = customPopulateFirstRow To customPopulateLastRow
        For y = customPopulateFirstColumn To customPopulateLastColumn
            If Cells(x, y).Value <> "" Then
                customDasgipArray(customTableCounter) = Cells(x, customPopulateFirstColumn - 1).Value
                customTimePointArray(customTableCounter) = Cells(x, y).Value
                customColorArray(customTableCounter) = Cells(x, customPopulateLastColumn + 1).Value
                customTableCounter = customTableCounter + 1
            End If
        Next
    Next
   
    'Clear table
    Call Clear_Table
   
    'Loop over vessels and timepoints
    For x = 1 To customTableCounter - 1
        'Check if plate needs to be wrapped to fourth row
        If (counter = 12) Or (counter = 10 And emptyRightTwoColumns) Then
            If Not wrapped Then
                firstRow = firstRow + 3
                counter = 0
                wrapped = True
            'Error notification for excessive samples
            Else
                MsgBox "Plate Overflow! Please adjust Custom Populate settings."
                'Break loop if overflowed
                Exit For
            End If
        End If

        'Populate timepoints
        For Z = 1 To 3
            'Vessel
            Range("C" & firstRow + Z + (8 * counter)).Value = customDasgipArray(x)
            'Timepoint
            Range("D" & firstRow + Z + (8 * counter)).Value = customTimePointArray(x)
            'Color
            Range("E" & firstRow + Z + (8 * counter)).Value = customColorArray(x)
        Next
        counter = counter + 1
    Next
    
    Call Generate_Plate
End Sub

'Clear pre-existing data
Sub Clear_Table()
    Range("C2:F" & Rows.Count).ClearContents
End Sub

Sub Clear_Plate()

    Dim obj As Object
    
    For Each obj In ActiveSheet.Shapes
        If obj.Name Like "*" & "Copy" Then
            obj.Delete
        End If
    Next obj
    
    Range("I1").Value = ""

End Sub
