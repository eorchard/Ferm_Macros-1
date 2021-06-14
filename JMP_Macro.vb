

'Function determines number of DG units based on number of data sheets

Function countDataSheets(countFromWorkbook) As Integer

    Dim i As Integer

    Dim xCount As Integer

   

    'Count number of sheets containing keyword 'Data'

    For i = 1 To countFromWorkbook.Sheets.Count

        If InStr(1, countFromWorkbook.Sheets(i).Name, "Data") > 0 Then xCount = xCount + 1

    Next

   

    'Minus 1 since the DASGIP raw data file has an extra blank sheet named 'Data'

    countDataSheets = xCount - 1

End Function

 

'Function removes timepoints before inoculation time, populates a new column with name of DG unit

Private Sub compressData(numberOfDataSheets, DG_Unit)

    For i = 1 To numberOfDataSheets

        Sheets("Data" & i).Select

       

        'Need this condition to check if Data Sheet # matches actual DG unit since the vessels aren't always run in sequence

        If Range("E1").Value Like "*" & i & "*" Then

            Columns("C:C").Select

            Selection.End(xlDown).Select

            Range("A2:A" & ActiveCell.Row - 1).EntireRow.Delete

            ActiveWindow.ScrollRow = 1

            Range("C1").EntireColumn.Insert

            Range("C1").Value = "DG"

            Range("C2:C" & Range("B2").End(xlDown).Row).Value = DG_Unit & i

           

        'If DG vessel doesn't match up with Data Sheet #, compare the other numbers

        Else

            For j = 1 To 8

                If Range("E1").Value Like "*" & j & "*" Then

                    Columns("C:C").Select

                    Selection.End(xlDown).Select

                    Range("A2:A" & ActiveCell.Row - 1).EntireRow.Delete

                    ActiveWindow.ScrollRow = 1

                    Range("C1").EntireColumn.Insert

                    Range("C1").Value = "DG"

                    Range("C2:C" & Range("B2").End(xlDown).Row).Value = DG_Unit & j

                End If

            Next

        End If

     Next

End Sub

 

'Function will consolidate data onto one sheet

Private Sub consolidateData(numberOfDataSheets)

    If (numberOfDataSheets > 1) Then

        For i = 2 To numberOfDataSheets

            Sheets("Data" & i).Select

            Range("A2:AO" & Range("B2").End(xlDown).Row).Copy

            Sheets("Data1").Select

            Columns("A:A").Select

            Selection.End(xlDown).Offset(1, 0).Select

            ActiveSheet.Paste

        Next

    End If

   

    'Remove number "1" from headers

    Worksheets("Data1").Rows("1").Replace What:="1", Replacement:=""

End Sub

 

'Function imports raw data file from DG units

Private Sub importRawData()

    Dim filter As String

    Dim rawDataFilename As String

    Dim rawDataWorkbook As Workbook

    Dim targetWorkbook As Workbook

    Dim DG_Unit As String

   

    filter = "Text files (*.xlsx),*.xlsx"

    MsgBox "Please select the DASGIP raw data file", vbOKOnly

   

    'JMP Macro workbook is the target

    Set targetWorkbook = Application.ThisWorkbook

   

    'Get raw data workbook

    rawDataFilename = Application.GetOpenFilename(filter, , caption)

    Set rawDataWorkbook = Application.Workbooks.Open(rawDataFilename)

   

    'This function makes the macro compatible with any number of DG vessels

    numberOfDataSheets = countDataSheets(rawDataWorkbook)

   

    'Identify which DASGIP the raw data is coming from, currently relies on filename

    If rawDataWorkbook.Name Like "*" & "DG3" & "*" Then

        DG_Unit = "DG3_u"

    ElseIf rawDataWorkbook.Name Like "*" & "DG4" & "*" Then

        DG_Unit = "DG4_u"

    ElseIf rawDataWorkbook.Name Like "*" & "DG5" & "*" Then

        DG_Unit = "DG5_u"

    End If

   

    'Remove timepoints before inoculation from raw data

    Call compressData(numberOfDataSheets, DG_Unit)

   

    'Copy data from DG raw files to JMP Macro

    For i = 1 To numberOfDataSheets

        Dim targetSheet As Worksheet

        Set targetSheet = targetWorkbook.Worksheets("Data" & i)

        Dim sourceSheet As Worksheet

        Set sourceSheet = rawDataWorkbook.Worksheets("Data" & i)

       

        'Identify last row in order to extract the correct range

        Dim lastRow As Long

        lastRow = Application.WorksheetFunction.CountA(Columns(1))

       

        targetSheet.Range("A2", "AI" & lastRow).Value = sourceSheet.Range("A2", "AI" & lastRow).Value

    Next

   

    'Close raw data file

    rawDataWorkbook.Close SaveChanges:=False

   

    'Append all DG raw data to bottom of first sheet

    Call consolidateData(numberOfDataSheets)

   

    'Convert Duration and InoculationTime to "hh:mm:ss" format

    lastRow = Application.WorksheetFunction.CountA(Columns(1))

   

    'Duration

    For Each cell In Range("B2:B" & lastRow)

        cell.Value = "=TEXT(""" & cell.Value & """, ""hh:mm:ss"")"

    Next cell

   

    'InoculationTime

    For Each cell In Range("D2:D" & lastRow)

        cell.Value = "=TEXT(""" & cell.Value & """, ""hh:mm:ss"")"

    Next cell

   

End Sub

 

'Main macro container

Sub Run_JMP_Macro()

    'Clear pre-existing data

    For i = 1 To 8

        Worksheets(i).Range("A2:AI" & Rows.Count).ClearContents

    Next

   

    'Import DG raw data file

    Call importRawData

End Sub

