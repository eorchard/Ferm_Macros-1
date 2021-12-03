'Function counts number of DO spikes based on Pump A totalizer values
Function countDOSpikes(lastRow, rawDataSheet) As Integer
    Dim numberOfSpikes As Integer
    Dim highTotalizer As Double, potentialHighTotalizer As Double
    For Each Cell In rawDataSheet.Range("AF2:AF" & lastRow)
        'Skip empty cells
        If (Cell.Value <> "") Then
 
            'Identifying DO spikes will require at least a 10mL increase
            If (Cell.Value > highTotalizer + 10) Then
 
                'potentialHighTotalizer becomes highTotalizer if new totalizer value appears in raw data 3x consecutively
                If (Cell.Value = potentialHighTotalizer) Then
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
                    potentialHighTotalizer = Cell.Value
                End If
            End If
        End If
    Next
    countDOSpikes = numberOfSpikes
End Function
 
'Sub adds four blank rows between each data point
Private Sub insertBlankRows(targetSheet, lastRow)
    Dim myRange As Range
    Dim counter&
    Dim timeDifference As Double, timeDifferenceMultipleOfFiveMinutes As Double
    Set myRange = targetSheet.Range("A2:M" & lastRow)

    'Attempt to improve performance since this sub is time-consuming
    ActiveSheet.DisplayPageBreaks = False
    With Application
        .ScreenUpdating = False
        .DisplayStatusBar = False
        .Calculation = xlCalculationManual
        .EnableEvents = False
    End With

    'Space out rows based on how long the time intervals are
    For i = (lastRow - 1) To 2 Step -1
        timeDifference = Round((myRange.Range("A" & i).Value - myRange.Range("A" & i - 1).Value), 4)
        timeDifferenceMultipleOfFiveMinutes = Round((timeDifference / (Round(5 / (60 * 24), 4))), 0) - 1

        If timeDifferenceMultipleOfFiveMinutes > 0 Then
            myRange.Rows(i).EntireRow.Resize(timeDifferenceMultipleOfFiveMinutes).Insert Shift:=xlDown
        ElseIf timeDifferenceMultipleOfFiveMinutes = -1 Then
            myRange.Rows(i).EntireRow.Delete
        End If
    Next

    For counter = myRange.Rows.Count To 2 Step -1
         myRange.Rows(counter).EntireRow.Offset(1).Resize(4).Insert Shift:=xlDown
    Next counter

   ActiveSheet.DisplayPageBreaks = True
   With Application
        .ScreenUpdating = True
        .DisplayStatusBar = True
        .Calculation = xlCalculationManual
        .EnableEvents = True
    End With
End Sub

'Sub adds DG Column to raw data
Private Sub addDGColumn(DG_Unit, vesselNumber, rawDataWorkbook, targetWorkbook)
    Dim strainID As String, customHeader1 As String, customHeader2 As String, customHeaderData1 As String, customHeaderData2 As String

    With targetWorkbook.Worksheets("Cover Sheet")
        customHeader1 = .Range("F3").Value
        customHeader2 = .Range("G3").Value

        If DG_Unit = "DG3_u" Then
            strainID = .Range("E" & 3 + vesselNumber).Value
            customHeaderData1 = .Range("F" & 3 + vesselNumber).Value
            Debug.Print .Range("F" & 3 + vesselNumber).Value
            customHeaderData2 = .Range("G" & 3 + vesselNumber).Value
        ElseIf DG_Unit = "DG4_u" Then
            strainID = .Range("E" & 11 + vesselNumber).Value
            customHeaderData1 = .Range("F" & 11 + vesselNumber).Value
            customHeaderData2 = .Range("G" & 11 + vesselNumber).Value
        ElseIf DG_Unit = "DG5_u" Then
            strainID = .Range("E" & 19 + vesselNumber).Value
            customHeaderData1 = .Range("F" & 19 + vesselNumber).Value
            customHeaderData2 = .Range("G" & 19 + vesselNumber).Value
        Else
            strainID = "N/A"
            customHeaderData1 = "N/A"
            customHeaderData2 = "N/A"
        End If
    End With

    ActiveWindow.ScrollRow = 1

    With Range("C1")
        If .Value <> "DG" Then
            .EntireColumn.Insert
            .Value = "DG"
        End If
    End With

    If Range("AJ1").Value <> "Strain ID" Then
        Range("AK1").EntireColumn.Insert
        Range("AJ1").Value = "Strain ID"
    End If

    'Name custom headers
    With targetWorkbook.Worksheets("Data1")
        .Range("AK1").Value = customHeader1
        .Range("AL1").Value = customHeader2
    End With

    'Fill columns
    Range("C2:C" & Range("B2").End(xlDown).Row).Value = DG_Unit & vesselNumber
    Range("AJ2:AJ" & Range("B2").End(xlDown).Row).Value = strainID
    Range("AK2:AK" & Range("B2").End(xlDown).Row).Value = customHeaderData1
    Range("AL2:AL" & Range("B2").End(xlDown).Row).Value = customHeaderData2
End Sub
 
'Sub populates columns needed for OUR/CER/RQ calculations
Private Sub fillColumnFromAbove(lastRow)
    'Turn off screen updating to improve performance
    Application.ScreenUpdating = False
    On Error Resume Next

    'Fill down F.PV
    With Columns("H")
        Range("H10", "H" & lastRow).SpecialCells(xlCellTypeBlanks).Formula = "=R[-1]C"
        .Value = .Value
    End With

    'Fill down T.PV
    With Columns("AC")
        Range("AC10", "AC" & lastRow).SpecialCells(xlCellTypeBlanks).Formula = "=R[-1]C"
        .Value = .Value
    End With

    'Fill down V.VPV
    With Columns("AE")
        Range("AE10", "AE" & lastRow).SpecialCells(xlCellTypeBlanks).Formula = "=R[-1]C"
        .Value = .Value
    End With

    'Fill down XCO2.PV
    With Columns("AH")
        Range("AH10", "AH" & lastRow).SpecialCells(xlCellTypeBlanks).Formula = "=R[-1]C"
        .Value = .Value
    End With

    'Fill down XO2.PV
    With Columns("AI")
        Range("AI10", "AI" & lastRow).SpecialCells(xlCellTypeBlanks).Formula = "=R[-1]C"
        .Value = .Value
    End With

    Err.Clear
    Application.ScreenUpdating = True
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

'Sub removes timepoints before inoculation time, populates a new column with name of DG unit
Private Sub removePreinoculationData(numberOfDataSheets, DG_Unit, rawDataWorkbook, targetWorkbook)
    For i = 1 To numberOfDataSheets
        Sheets("Data" & i).Select

        'Need this condition to check if Data Sheet # matches actual DG unit since the vessels aren't always run in sequence
        If Range("E1").Value Like "*" & i & "*" Then
            Columns("C:C").Select
            Selection.End(xlDown).Select
            Range("A2:A" & ActiveCell.Row - 1).EntireRow.Delete
            Call addDGColumn(DG_Unit, i, rawDataWorkbook, targetWorkbook)

        'If DG vessel doesn't match up with Data Sheet #, compare the other numbers
        Else
            For j = 1 To 8
                If Range("E1").Value Like "*" & j & "*" Then
                    Columns("C:C").Select
                    Selection.End(xlDown).Select
                    Range("A2:A" & ActiveCell.Row - 1).EntireRow.Delete
                    Call addDGColumn(DG_Unit, j, rawDataWorkbook, targetWorkbook)
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
            If Range("A2").Value <> "" Then
                Range("A2:AY" & Range("B2").End(xlDown).Row).Copy
                Sheets("Data1").Select
                Columns("A:A").Select
                Selection.End(xlDown).Offset(1, 0).Select
                ActiveSheet.Paste
            End If
        Next
    End If

    'Round each timepoint, makes time synchronization easier
    lastRow = Application.WorksheetFunction.CountA(Columns(1))
    timeArray = Worksheets("Data1").Range("A2:A" & lastRow).Value
 
    For j = 1 To UBound(timeArray, 1)
        timeArray(j, 1) = "=MROUND(""" & timeArray(j, 1) & """, ""0:01"")"
    Next

    Worksheets("Data1").Range("A2:A" & lastRow).Value = timeArray
 
    'Remove number "1" from headers
    Worksheets("Data1").Range("A1:AI1").Replace What:="1", Replacement:=""
End Sub
 
'Sub will import OUR data
Private Sub importOURData(dasgipRawDataFileName)
    Dim rawDataWorkbook As Workbook, targetWorkbook As Workbook
    Dim rawDataSheet As Worksheet, targetSheet As Worksheet
    Dim yy As String, mm As String, dd As String, ddOriginal As String, filter As String, fileFound As String, vesselID As String
    Dim numberOfDaysPerMonthArray As Variant
    Dim lastRowTarget As Integer, lastRowRaw As Integer, lastRowOUR As Integer, matchedRowNumberOUR As Integer, matchedRowNumberDG As Integer, _
        dasgipID As Integer
    Dim hasAnotherDataFile As Boolean, hasExistingData As Boolean, exitFor As Boolean
    Dim datePattern As Object, datePatternRegExp As Object, dasgipIDPattern As Object, dasgipIDRegExp As Object
    Dim firstTimeToMatch As Double

    numberOfDaysPerMonthArray = Array(31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31)
    Set datePatternRegExp = New RegExp
    Set dasgipIDRegExp = New RegExp
    fileNotFound = True
 
    'Get OUR raw data
    filter = "Text files (*.xlsx),*.xlsx"
    Set targetWorkbook = Application.ThisWorkbook
 
    'Parse OUR date to query based off of date from DG raw data filename
    datePatternRegExp.Pattern = "\d{6}"
    Set datePattern = datePatternRegExp.Execute(dasgipRawDataFileName)
    dd = Right(datePattern(0), 2)
    mm = Mid(datePattern(0), 3, 2)
    yy = Left(datePattern(0), 2)
    ddOriginal = dd
    mmOriginal = mm

    'Pattern must include "raw data\" to exclude false matches from the batch ID
    dasgipIDRegExp.Pattern = "raw data\\.*DG\d"
   Set dasgipIDPattern = dasgipIDRegExp.Execute(dasgipRawDataFileName)
    dasgipID = Right(dasgipIDPattern(0), 1)

    'Copy paste data from each OUR file into sheet tab (OUR1, OUR2, etc..)
    For i = 1 To 8
        hasExistingData = False
 
        'Parse cell C2 to get the correct DG unit; can't rely on worksheet name. Use vessel number to open raw data file
        If Worksheets("Data" & i).Cells(2, 3) <> "" Then
            vesselID = Right(Worksheets("Data" & i).Cells(2, 3), 1)
            fileFound = Dir("S:\Projects\Fermentation\Ferm&StrainDevelopment\OUR Data\20" & yy & "\" & dasgipID & _
                "-" & vesselID & "\analysis\" & mm & dd & "*.csv")
            Set targetSheet = targetWorkbook.Worksheets("OUR" & vesselID)
        End If

        'Run block if file exists
        If fileFound <> "" Then
            fileNotFound = False

            'Collect OUR data for individual DG unit until no more sequential data files exist
            Do
                Set rawDataWorkbook = Application.Workbooks.Open("S:\Projects\Fermentation\Ferm&StrainDevelopment\OUR Data\20" & yy & "\" & dasgipID & _
                    "-" & vesselID & "\analysis\" & mm & dd & "*.csv")
                Set rawDataSheet = rawDataWorkbook.Worksheets(1)
 
                'Identify last row in order to extract the correct range
                lastRowTarget = targetSheet.Range("A" & Rows.Count).End(xlUp).Row
                lastRowRaw = rawDataSheet.Range("A" & Rows.Count).End(xlUp).Row
 
                If Not hasExistingData Then
                    'Copy data from OUR raw files to JMP Macro, no appending logic needed for the first day
                    targetSheet.Range("A1", "M" & lastRowRaw).Value = rawDataSheet.Range("A1", "M" & lastRowRaw).Value
                    hasExistingData = True
                Else
                    'Append to bottom if not the first day
                    lastRowTarget = lastRowTarget + 1
                    targetSheet.Range("A" & lastRowTarget, "M" & lastRowTarget + lastRowRaw - 2).Value = rawDataSheet.Range("A2", "M" & lastRowRaw).Value
                End If
 
                'Increment day
                If CInt(dd) >= 9 Then
                    dd = CStr(CInt(dd) + 1)
                ElseIf CInt(dd) < 9 Then
                    dd = CStr("0" & CInt(dd) + 1)
                End If

                'Cycle dd, mm, or yy when they exceed the max value
                If (CInt(dd) > numberOfDaysPerMonthArray(CInt(mm) - 1)) Then
                    dd = "01"
                    If mm <> "09" Then
                        mm = IIf(mm = "12", "01", CStr("0" & CInt(mm + 1)))
                    ElseIf mm = "09" Then
                        mm = "10"
                    ElseIf mm = "13" Then
                        yy = CStr(CInt(yy) + 1)
                        mm = "01"
                    End If
                End If
 
                'Check if an OUR data file for the next day exists
                fileFound = Dir("S:\Projects\Fermentation\Ferm&StrainDevelopment\OUR Data\20" & yy & "\" & dasgipID & _
                    "-" & vesselID & "\analysis\" & mm & dd & "*.csv")
                hasAnotherDataFile = IIf(fileFound <> "", True, False)
 
                'Set dd and mm back to normal when all unit's data files are scraped
                If Not hasAnotherDataFile Then
                    dd = ddOriginal
                    mm = mmOriginal

                    With targetSheet
                      'When no more data files exist, round all the OUR timepoints to nearest fifth minute. Makes time synchronization easier
                      lastRowOUR = .Cells(Rows.Count, 1).End(xlUp).Row
                      timeArrayOUR = .Range("A2:A" & lastRowOUR).Value

                      For j = 1 To UBound(timeArrayOUR, 1)
                          timeArrayOUR(j, 1) = "=MROUND(""" & timeArrayOUR(j, 1) & """, ""0:05"")"
                      Next

                      .Range("A2:A" & lastRowOUR).Value = timeArrayOUR
                    End With
                End If
 
                rawDataWorkbook.Close SaveChanges:=False
 
            Loop While hasAnotherDataFile = True

            'Determine data output intervals then space them out by this amount
            Call insertBlankRows(targetSheet, lastRowOUR)

            'Synchronize DG timepoints and OUR timepoints
            exitFor = False

            'Update lastRowOUR after inserting blank rows
            lastRowOUR = targetSheet.Cells(Rows.Count, 1).End(xlUp).Row

            For j = 2 To Rows.Count
                '1440 min/day
                'Skip first OUR row, sometimes the first data timepoint is off
                For k = 3 To 1440 Step 5
                    If targetWorkbook.Worksheets("Data" & i).Cells(j, 1).Value = targetWorkbook.Worksheets("OUR" & vesselID).Cells(k, 1).Value Then
                        'Matching timepoints
                        matchedRowNumberOUR = k
                        matchedRowNumberDG = j

                        'Copy over matching OUR data to the corresponding row in the DG data
                        targetWorkbook.Worksheets("OUR" & vesselID).Range("A" & matchedRowNumberOUR & ":N" & lastRowOUR).Copy _
                            Worksheets("Data" & i).Range("AM" & matchedRowNumberDG & ":BA" & (matchedRowNumberDG + lastRowOUR))

                        'Break inner loop
                        exitFor = True
                        Exit For
                    End If
                Next
                'Break outer loop
                If exitFor Then Exit For
            Next
            
        'Display MsgBox if no OUR data files are found
        ElseIf fileNotFound And i = 8 Then
            MsgBox "No OUR data files exist for " & yy & "-" & mm & "-" & dd
        End If
    Next
End Sub

'Function imports raw data file from DG units
Private Sub importRawData()
    Dim filter As String, DG_Unit As String, columnAr As String, columnXO2 As String, columnOUR As String, _
        columnCER As String, columnRQ As String, columnOURTime As String, columnN2 As String, columnTPV As String, _
        columnVPV As String, columnXCO2 As String, columnCustomHeader1 As String, columnCustomHeader2 As String, _
        columnO2 As String, columnFPV As String
    Dim rawDataSheet As Worksheet, targetSheet As Worksheet
    Dim rawDataWorkbook As Workbook, targetWorkbook As Workbook
    Dim numberOfSpikes As Integer, lastRow As Integer, answer As Integer, totalRowCount As Integer
    Dim importOUR As Boolean, isArgonLow As Boolean

    isArgonLow = False
    columnFPV = "H"
    columnTPV = "AC"
    columnVPV = "AE"
    columnXCO2 = "AH"
    columnXO2 = "AI"
    columnCustomHeader1 = "AK"
    columnCustomHeader2 = "AL"
    columnOURTime = "AM"
    columnN2 = "AN"
    columnO2 = "AO"
    columnAr = "AP"
    columnCO2 = "AQ"
    columnOUR = "AZ"
    columnCER = "BA"
    columnRQ = "BB"
    filter = "Text files (*.xlsx),*.xlsx"
    MsgBox "Please select the DASGIP raw data file", vbOKOnly

    'JMP Macro workbook is the target
    Set targetWorkbook = Application.ThisWorkbook

    'Get raw data workbook
    rawDataFileName = Application.GetOpenFilename(filter, , Caption)

    If rawDataFileName = False Then
        Exit Sub
    Else
        Set rawDataWorkbook = Application.Workbooks.Open(rawDataFileName)
    End If

    'This function makes the macro compatible with any number of DG vessels
    numberOfDataSheets = countDataSheets(rawDataWorkbook)

    With rawDataWorkbook
    'Identify which DASGIP the raw data is coming from, currently relies on filename
        If .Name Like "*" & "DG3" & "*" Then
            DG_Unit = "DG3_u"
        ElseIf .Name Like "*" & "DG4" & "*" Then
            DG_Unit = "DG4_u"
        ElseIf .Name Like "*" & "DG5" & "*" Then
            DG_Unit = "DG5_u"
        End If
    End With

    'Remove timepoints before inoculation from raw data
    Call removePreinoculationData(numberOfDataSheets, DG_Unit, rawDataWorkbook, targetWorkbook)
   
    'Copy data from DG raw files to JMP Macro
    For i = 1 To numberOfDataSheets
        Set targetSheet = targetWorkbook.Worksheets("Data" & i)
        Set rawDataSheet = rawDataWorkbook.Worksheets("Data" & i)
       
        'Identify last row in order to extract the correct range
        lastRow = Application.WorksheetFunction.CountA(rawDataSheet.Range("A:A"))
       
        'DG3 and DG5 raw data export contain 6 additional columns than DG4, remove these columns so all DG units are formatted the same way
        'Update 02DEC2021 - This may  not be needed anymore
        If Application.WorksheetFunction.CountA(rawDataSheet.Range("AN:AN")) <> 0 Then
            rawDataSheet.Range("J:J,P:P,R:R,T:T,AL:AL,AN:AN").Delete
        End If
       
        With targetSheet
            'Import data
            .Range("A2", columnCustomHeader2 & lastRow).Value = rawDataSheet.Range("A2", columnCustomHeader2 & lastRow).Value
            
            If .Range("A2").Value <> "" Then
                'Round each timepoint, makes time synchronization easier
                timeArray = .Range("A2:" & columnCustomHeader2 & lastRow).Value
            
                For j = 1 To UBound(timeArray, 1)
                    timeArray(j, 1) = "=MROUND(""" & timeArray(j, 1) & """, ""0:01"")"
                Next
              
                .Range("A2:A" & lastRow).Value = timeArray
               
                'Get number of DO spikes
                numberOfSpikes = countDOSpikes(lastRow, rawDataSheet)
            End If
        End With
    Next
   
    'Close raw data file
    rawDataWorkbook.Close SaveChanges:=False
   
    'Import OUR data if selected
    answer = MsgBox("Would you like to import OUR data? (This may take a few minutes.)", vbYesNo)
    importOUR = IIf(answer = 6, True, False)
 
    If importOUR Then
        Call importOURData(rawDataFileName)
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
    
    'Final data processing if OUR was selected
    If importOUR Then
        'Insert OUR/CER/RQ formulas
        Dim gasFormulas(1 To 3) As Variant
        
        With Worksheets("Data1")
            gasFormulas(1) = "=(((" & columnFPV & "12/60)*(" & columnXO2 & "12-" & columnO2 & "12)*600)/((" & columnVPV & "12/1000)*0.082057*(" & columnTPV & "12+273.15)))"
            gasFormulas(2) = "=(((" & columnFPV & "12/60)*(" & columnCO2 & "12-" & columnXCO2 & "12)*600)/((" & columnVPV & "12/1000)*0.082057*(" & columnTPV & "12+273.15)))"
            gasFormulas(3) = "=" & columnCER & "12/" & columnOUR & "12"
    
            .Range(columnOUR & "12:" & columnRQ & "12").Formula = gasFormulas
            .Range(columnOUR & "12:" & columnRQ & lastRow).FillDown
        End With
        
        'Fill columns needed for OUR/CER/RQ calculations
        Call fillColumnFromAbove(lastRow)
        
        'Fix data types that tend to revert
        Range(columnCustomHeader1 & "2:" & columnCustomHeader2 & lastRow).NumberFormat = "General"
        Range(columnN2 & "2:" & columnN2 & lastRow).NumberFormat = "General"
        
        'Final step is to remove all excess data
        For i = 11 To lastRow
            If Range(columnOURTime & i).Value = "" Then
                Range(columnN2 & i & ":" & columnRQ & i).ClearContents
            End If
            If (Range(columnAr & i).Value <> "") Then
                If (Range(columnXO2 & i).Value = 21) And (Range(columnAr & i).Value < 0.85) Then
                    isArgonLow = True
                    Range(columnAr & i).Interior.ColorIndex = 3
                End If
            End If
        Next
        
        'Add to events table and color code if needed
        With Worksheets("Cover Sheet").Range("Q4")
            .Value = IIf(isArgonLow, "Low Argon levels detected", "Argon levels normal")
            .Interior.ColorIndex = IIf(isArgonLow, 3, 0)
        End With
        
        
        'Save workbook so the OUR/CER/RQ data shows up
        ActiveWorkbook.Save
    End If
End Sub

'Main macro container
Sub Run_JMP_Macro()
    'Clear pre-existing data except Cover Sheet
    For i = 2 To 17
        Worksheets(i).Range("A2:BB" & Rows.Count).ClearContents
    Next
 
    'Clear colors and custom headers
    With Worksheets("Data1")
        .Range("A2:BB" & Rows.Count).Interior.ColorIndex = 0
        .Range("AK1:AL1").ClearContents
    End With
    
    'Clear alerts
    Worksheets("Cover Sheet").Range("Q4:Q27").ClearContents
    
    'Import DG raw data file
    Call importRawData
End Sub

