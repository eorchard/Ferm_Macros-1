'Function counts number of DO spikes based on Pump A totalizer values
Function countDOSpikes(lastRow, rawDataSheet) As Integer
    Dim numberOfSpikes As Integer
    Dim highTotalizer As Double, potentialHighTotalizer As Double
 
    For Each cell In rawDataSheet.Range("AF2:AF" & lastRow)
        'Skip empty cells
        If (cell.Value <> "") Then

            'Identifying DO spikes will require at least a 10mL increase
            If (cell.Value > highTotalizer + 10) Then

                'potentialHighTotalizer becomes highTotalizer if new totalizer value appears in raw data 3x consecutively
                If (cell.Value = potentialHighTotalizer) Then
                    If counter = 3 Then
                        highTotalizer = potentialHighTotalizer
                        counter = 0
                        numberOfSpikes = numberOfSpikes + 1

                    'Increment counter until 3
                    Else
                        counter = counter + 1
                    End If

                'New high value
                Else
                    potentialHighTotalizer = cell.Value
                End If
            End If
        End If
    Next
    countDOSpikes = numberOfSpikes
End Function
 
'Sub adds DG Column to raw data
Private Sub addDGColumn(DG_Unit, vesselNumber)
    Columns("C:C").Select
    Selection.End(xlDown).Select
    Range("A2:A" & ActiveCell.Row - 1).EntireRow.Delete
    ActiveWindow.ScrollRow = 1
    Range("C1").EntireColumn.Insert
    Range("C1").Value = "DG"
    Range("C2:C" & Range("B2").End(xlDown).Row).Value = DG_Unit & vesselNumber
End Sub
 
'Function determines number of DG units based on number of data sheets
Function countDataSheets(countFromWorkbook) As Integer
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
            Call addDGColumn(DG_Unit, i)
           
        'If DG vessel doesn't match up with Data Sheet #, compare the other numbers
        Else
            For j = 1 To 8
                If Range("E1").Value Like "*" & j & "*" Then
                    Call addDGColumn(DG_Unit, j)
                End If
            Next
        End If
     Next
End Sub
 
'Sub will consolidate data onto one sheet
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

'Sub will import OUR data
Private Sub importOURData()
    Dim rawDataWorkbook As Workbook, targetWorkbook As Workbook
    Dim rawDataSheet As Worksheet, targetSheet As Worksheet
    Dim rawDataFileName As String, mm As String, dd As String, ddOriginal As String, filter As String, fileFound As String
    Dim numberOfDaysPerMonthArray As Variant
    Dim lastRow As Integer
    Dim hasAnotherDataFile As Boolean, hasExistingData As Boolean
    
    numberOfDaysPerMonthArray = Array(31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31)

    MsgBox "Please select the first OUR raw data file", vbOKOnly

    'Get OUR raw data
    filter = "Text files (*.xlsx),*.xlsx"
    rawDataFileNameFull = Application.GetOpenFilename(filter, , caption)
    Set targetWorkbook = Application.ThisWorkbook

    'Parse date
    rawDataFileNameFull = Right(rawDataFileNameFull, 12)
    rawDataFileName = Left(rawDataFileNameFull, 8)
    dd = Left(rawDataFileName, 4)
    dd = Right(dd, 2)
    ddOriginal = dd
    mm = Left(rawDataFileName, 2)

    'Copy paste data from each OUR file into hidden sheet tab (OUR1, OUR2, etc..)
    For i = 1 To 8
        'TODO: Iterate over each workbook, copy paste data. Clear data on each OUR1-8 sheet before pasting data
        hasExistingData = False
        fileFound = Dir("S:\Projects\Fermentation\Ferm&StrainDevelopment\OUR Data\FOUR-" & i & "\analysis\" & mm & dd & "*.csv")

        If fileFound <> "" Then
            'Collect OUR data for individual DG unit until no more sequential data files exist
            Do
                Set rawDataWorkbook = Application.Workbooks.Open("S:\Projects\Fermentation\Ferm&StrainDevelopment\OUR Data\FOUR-" & i & "\analysis\" & mm & dd & "*.csv")
                Set rawDataSheet = rawDataWorkbook.Worksheets(1)
                Set targetSheet = targetWorkbook.Worksheets("OUR" & i)

                'Identify last row in order to extract the correct range
                lastRow = Application.WorksheetFunction.CountA(Columns(1))

                If Not hasExistingData
                    'Copy data from OUR raw files to JMP Macro, only for first day
                    targetSheet.Range("A2", "M" & lastRow).Value = rawDataSheet.Range("A2", "M" & lastRow).Value
                Else
                    rawDataSheet.Range("A2:M" & Range("B2").End(xlDown).Row).Copy
                    Sheets("OUR" & i).Select
                    Columns("A:A").Select
                    Selection.End(xlDown).Offset(1, 0).Select
                    ActiveSheet.Paste
                End If

                'Increment day
                If CInt(dd) >= 9
                    dd = CStr(CInt(dd) + 1)
                Else If CInt(dd) < 9
                    dd = CStr("0" & CInt(dd) + 1)
                
                'Reset dd or mm when they exceed the max value
                Else If (CInt(dd) >= numberOfDaysPerMonthArray(CInt(mm) - 1))
                    dd = "01"
                    IIf (mm = 12, mm = 1, mm + 1)
                End If

                'Check if an OUR data file for the next day exists
                fileFound = Dir("S:\Projects\Fermentation\Ferm&StrainDevelopment\OUR Data\FOUR-" & i & "\analysis\" & mm & dd & "*.csv")
                hasAnotherDataFile = IIf(fileFound <> "", True, False) 

                If Not hasAnotherDataFile Then
                    dd = ddOriginal
                End If

                rawDataWorkbook.Close SaveChanges:=False

            Loop While hasAnotherDataFile = True
        End If
    Next


    'Close raw data file
    rawDataWorkbook.Close SaveChanges:=False
End Sub
 
'Function imports raw data file from DG units
Private Sub importRawData()
    Dim filter As String, DG_Unit As String
    Dim rawDataSheet As Worksheet, targetSheet As Worksheet
    Dim rawDataWorkbook As Workbook, targetWorkbook As Workbook
    Dim numberOfSpikes As Integer, lastRow As Integer, answer As Integer
    Dim importOUR As Boolean
   
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
        Set targetSheet = targetWorkbook.Worksheets("Data" & i)
        Set rawDataSheet = rawDataWorkbook.Worksheets("Data" & i)
       
        'Identify last row in order to extract the correct range
        lastRow = Application.WorksheetFunction.CountA(Columns(1))
       
        'DG3 and DG5 raw data export contain 6 additional columns than DG4, remove these columns so all DG units are formatted the same way
        If Application.WorksheetFunction.CountA(rawDataSheet.Range("AN:AN")) <> 0 Then
            rawDataSheet.Range("J:J,P:P,R:R,T:T,AL:AL,AN:AN").Delete
        End If
       
        targetSheet.Range("A2", "AI" & lastRow).Value = rawDataSheet.Range("A2", "AI" & lastRow).Value
       
        numberOfSpikes = countDOSpikes(lastRow, rawDataSheet)
    Next
   
    'Close raw data file
    rawDataWorkbook.Close SaveChanges:=False
   
    'Import OUR data if selected
    answer = MsgBox("Would you like to import OUR data?", vbYesNo)
    importOUR = IIf(answer = 6, True, False)

    If importOUR Then
        Call importOURData()
    End If

    'Append all DG raw data to bottom of first sheet
    Call consolidateData(numberOfDataSheets)
   
 'Convert Duration to array, perform "[h]:mm:ss" conversion, insert back into spreadsheet
    lastRow = Application.WorksheetFunction.CountA(Columns(1))
    timeArray = Worksheets("Data1").Range("B2:B" & lastRow).Value
 
    For i = 1 To UBound(timeArray, 1)
        timeArray(i, 1) = "=TEXT(""" & timeArray(i, 1) & """, ""[h]:mm:ss"")"
    Next
   
    With Worksheets("Data1")
        .Range("B2:B" & lastRow).Value = timeArray
    End With
   
    'Convert InoculationTime to array, perform text conversion, insert back into spreadsheet
    timeArray = Worksheets("Data1").Range("D2:D" & lastRow).Value
 
    For i = 1 To UBound(timeArray, 1)
        timeArray(i, 1) = "=TEXT(""" & timeArray(i, 1) & """, ""[h]:mm:ss"")"
    Next
   
    With Worksheets("Data1")
       .Range("D2:D" & lastRow).Value = timeArray
    End With
End Sub
 
'Main macro container
Sub Run_JMP_Macro()
    'Clear pre-existing data
    For i = 1 To 8
        Worksheets(i).Range("A2:AV" & Rows.Count).ClearContents
    Next
   
    'Import DG raw data file
    Call importRawData
End Sub
