Sub Generate_Plate()

 

Dim ovalShape As Shape

Dim A1_X As Double

Dim A1_Y As Double

 

'Absolute location of well A1, rest will be calculated relatively

A1_X = 397.5

A1_Y = 60.75

 

'Initialize plate map in correct location with correct size

With ActiveSheet.Shapes("Plate Map")

    .Left = 366

    .Top = 0

    .Height = 378.72

    .Width = 510.48

End With

 

'Populate arrays with values, begin at x+1 since the first row is a header

Dim colorArray(96) As String

Dim timepointArray(96) As String

Dim dasgipArray(96) As String

Dim x As Integer

 

For x = 1 To 96

    dasgipArray(x) = Cells(x + 1, 3).Value

    timepointArray(x) = Cells(x + 1, 4).Value

    colorArray(x) = Cells(x + 1, 5).Value

Next

   

'Populate wells

Dim color As String

Dim counter As Integer

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

                    .Characters.Text = dasgipArray(counter) & vbNewLine & timepointArray(counter)

                    .MarginLeft = 0

                    .MarginRight = 0

                    .MarginTop = 0

                    .MarginBottom = 0

                    .Characters.Font.Size = 9

                    .HorizontalAlignment = xlHAlignCenter

                End With

            End With

        End If

        counter = counter + 1

    Next

Next

   

End Sub
