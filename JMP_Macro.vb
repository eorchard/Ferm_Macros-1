'Function counts number of DO spikes based on Pump A totalizer values
Function countDOSpikes(lastRow, sourceSheet) As Integer
    Dim numberOfSpikes As Integer
    Dim highTotalizer As Double
    Dim potentialHighTotalizer As Double
 
    numberOfSpikes = 0
    highTotalizer = 0
    potentialHighTotalizer = 0
 
    For Each cell In sourceSheet.Range("AF2:AF" & lastRow)
        'Skip empty cells
        If cell.Value <> "" Then
            'Identifying DO spikes will require at least a 10mL increase
            If (cell.Value > highTotalizer + 10) Then
                'potentialHighTotalizer becomes highTotalizer if new totalizer value appears in raw data 3x consecutively
                'This prevents the macro from identifying a gradual increase as multiple spikes
                If (cell.Value = potentialHighTotalizer) And (counter = 3) Then
                    highTotalizer = potentialHighTotalizer
                    counter = 0
                    numberOfSpikes = numberOfSpikes + 1
                   
                'Increment counter until 3
                ElseIf (cell.Value = potentialHighTotalizer) And (counter < 3) Then
                    counter = counter + 1
                   
                'New high value
                ElseIf cell.Value <> potentialHighTotalizer Then
                    potentialHighTotalizer = cell.Value
                    counter = counter + 1
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
    Dim numberOfSpikes As Integer
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
       
        'DG3 and DG5 raw data export contain 6 additional columns than DG4, remove these columns so all DG units are formatted the same way
        If Application.WorksheetFunction.CountA(sourceSheet.Range("AN:AN")) <> 0 Then
            sourceSheet.Range("J:J,P:P,R:R,T:T,AL:AL,AN:AN").Delete
        End If
       
        targetSheet.Range("A2", "AI" & lastRow).Value = sourceSheet.Range("A2", "AI" & lastRow).Value
       
        numberOfSpikes = countDOSpikes(lastRow, sourceSheet)
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
