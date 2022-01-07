'Function counts number of DO spikes based on Pump A totalizer values
Function countDOSpikes(lastRow, rawDataSheet, columnVAPV, rowDataStart) As Integer
    Dim numberOfSpikes As Integer
    Dim highTotalizer As Double, potentialHighTotalizer As Double
 
    For Each Cell In rawDataSheet.Range(columnVAPV & rowDataStart & ":" & columnVAPV & lastRow)
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
Private Sub insertBlankRows(targetSheet, lastColumnOUR, lastRowOUR, columnTimestamp, rowDataStart)
    Dim myRange As Range
    Dim counter&
    Dim timeDifference As Double, timeDifferenceMultipleOfFiveMinutes As Double
    Set myRange = targetSheet.Range(columnTimestamp & rowDataStart & ":" & lastColumnOUR & lastRowOUR)
    
    'Attempt to improve performance since this sub is time-consuming
    ActiveSheet.DisplayPageBreaks = False
    With Application
        .ScreenUpdating = False
        .DisplayStatusBar = False
        .Calculation = xlCalculationManual
        .EnableEvents = False
    End With
    
    'Space out rows based on how long the time intervals are
    For i = (lastRowOUR - 1) To 2 Step -1
        timeDifference = Round((myRange.Range(columnTimestamp & i).Value - myRange.Range(columnTimestamp & i - 1).Value), 4)
        
        'This calculation makes the script scalable
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
    
    With Application
        .ScreenUpdating = True
        .DisplayStatusBar = True
        .Calculation = xlCalculationManual
        .EnableEvents = True
    End With
End Sub
 
'Sub adds DG Column to raw data
Private Sub addDGColumn(DG_Unit, vesselNumber, rawDataWorkbook, targetWorkbook, columnDuration, columnDG, columnStrainIDInput, columnStrainID, _
        columnCustomHeader1Input, columnCustomHeader2Input, columnCustomHeader1, columnCustomHeader2, rowFirst, rowDataStart)
        
    Dim strainID As String, customHeader1 As String, customHeader2 As String, customHeaderData1 As String, customHeaderData2 As String
    Dim rowHeader As Integer
    
    rowHeader = 3

    With targetWorkbook.Worksheets("Cover Sheet")
        customHeader1 = .Range(columnCustomHeader1Input & rowHeader).Value
        customHeader2 = .Range(columnCustomHeader2Input & rowHeader).Value
        
        If DG_Unit = "DG3_u" Then
            strainID = .Range(columnStrainIDInput & rowHeader + vesselNumber).Value
            customHeaderData1 = .Range(columnCustomHeader1Input & rowHeader + vesselNumber).Value
            customHeaderData2 = .Range(columnCustomHeader2Input & rowHeader + vesselNumber).Value
        ElseIf DG_Unit = "DG4_u" Then
            strainID = .Range(columnStrainIDInput & (rowHeader + 8) + vesselNumber).Value
            customHeaderData1 = .Range(columnCustomHeader1Input & (rowHeader + 8) + vesselNumber).Value
            customHeaderData2 = .Range(columnCustomHeader2Input & (rowHeader + 8) + vesselNumber).Value
        ElseIf DG_Unit = "DG5_u" Then
            strainID = .Range(columnStrainIDInput & (rowHeader + 16) + vesselNumber).Value
            customHeaderData1 = .Range(columnCustomHeader1Input & (rowHeader + 16) + vesselNumber).Value
            customHeaderData2 = .Range(columnCustomHeader2Input & (rowHeader + 16) + vesselNumber).Value
        Else
            strainID = "N/A"
            customHeaderData1 = "N/A"
            customHeaderData2 = "N/A"
        End If
    End With
    
    ActiveWindow.ScrollRow = 1
    
    If Range(columnDG & rowFirst).Value <> "DG" Then
        Range(columnDG & rowFirst).EntireColumn.Insert
        Range(columnDG & rowFirst).Value = "DG"
    End If
    
    If Range(columnStrainID & rowFirst).Value <> "Strain ID" Then
        Range(columnCustomHeader1 & rowFirst).EntireColumn.Insert
        Range(columnStrainID & rowFirst).Value = "Strain ID"
    End If
    
    'Name custom headers
    With targetWorkbook.Worksheets("Data1")
        .Range(columnCustomHeader1 & rowFirst).Value = customHeader1
        .Range(columnCustomHeader2 & rowFirst).Value = customHeader2
    End With
    
    'Fill columns
    Range(columnDG & rowDataStart & ":" & columnDG & Range(columnDuration & rowDataStart).End(xlDown).Row).Value = DG_Unit & vesselNumber
    Range(columnStrainID & rowDataStart & ":" & columnStrainID & Range(columnDuration & rowDataStart).End(xlDown).Row).Value = strainID
    Range(columnCustomHeader1 & rowDataStart & ":" & columnCustomHeader1 & Range(columnDuration & rowDataStart).End(xlDown).Row).Value = customHeaderData1
    Range(columnCustomHeader2 & rowDataStart & ":" & columnCustomHeader2 & Range(columnDuration & rowDataStart).End(xlDown).Row).Value = customHeaderData2
End Sub

'Sub populates columns needed for OUR/CER/RQ calculations
Private Sub fillColumnFromAbove(lastRow, columnFPV, columnTPV, columnVPV, columnXCO2, columnXO2)
    'Turn off screen updating to improve performance
    Application.ScreenUpdating = False
    On Error Resume Next
    
    'Fill down F.PV
    With Columns(columnFPV)
        Range(columnFPV & "10", columnFPV & lastRow).SpecialCells(xlCellTypeBlanks).Formula = "=R[-1]C"
        .Value = .Value
    End With
    
    'Fill down T.PV
    With Columns(columnTPV)
        Range(columnTPV & "10", columnTPV & lastRow).SpecialCells(xlCellTypeBlanks).Formula = "=R[-1]C"
        .Value = .Value
    End With
    
    'Fill down V.VPV
    With Columns(columnVPV)
        Range(columnVPV & "10", columnVPV & lastRow).SpecialCells(xlCellTypeBlanks).Formula = "=R[-1]C"
        .Value = .Value
    End With
    
    'Fill down XCO2.PV
    With Columns(columnXCO2)
        Range(columnXCO2 & "10", columnXCO2 & lastRow).SpecialCells(xlCellTypeBlanks).Formula = "=R[-1]C"
        .Value = .Value
    End With
    
    'Fill down XO2.PV
    With Columns(columnXO2)
        Range(columnXO2 & "10", columnXO2 & lastRow).SpecialCells(xlCellTypeBlanks).Formula = "=R[-1]C"
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
Private Sub removePreinoculationData(numberOfDataSheets, DG_Unit, rawDataWorkbook, targetWorkbook, columnDuration, columnDG, columnStrainIDInput, columnStrainID, _
        columnCustomHeader1Input, columnCustomHeader2Input, columnCustomHeader1, columnCustomHeader2, columnFPV, rowHeader, rowDataStart)
    
    For i = 1 To numberOfDataSheets
        Sheets("Data" & i).Select
       
        'Need this condition to check if Data Sheet # matches actual DG unit since the vessels aren't always run in sequence
        If Range(columnFPV & rowHeader).Value Like "*" & i & "*" Then
            Columns("C:C").Select
            Selection.End(xlDown).Select
            Range("A2:A" & ActiveCell.Row - 1).EntireRow.Delete
            Call addDGColumn(DG_Unit, i, rawDataWorkbook, targetWorkbook, columnDuration, columnDG, columnStrainIDInput, columnStrainID, _
        columnCustomHeader1Input, columnCustomHeader2Input, columnCustomHeader1, columnCustomHeader2, rowHeader, rowDataStart)
           
        'If DG vessel doesn't match up with Data Sheet #, compare the other numbers
        Else
            For j = 1 To 8
                If Range(columnFPV & rowHeader).Value Like "*" & j & "*" Then
                    Columns("C:C").Select
                    Selection.End(xlDown).Select
                    Range("A2:A" & ActiveCell.Row - 1).EntireRow.Delete
                    Call addDGColumn(DG_Unit, j, rawDataWorkbook, targetWorkbook, columnDuration, columnDG, columnStrainIDInput, columnStrainID, _
        columnCustomHeader1Input, columnCustomHeader2Input, columnCustomHeader1, columnCustomHeader2, rowHeader, rowDataStart)
                End If
            Next
        End If
     Next
End Sub
 
'Sub will consolidate data onto one sheet
Private Sub consolidateData(numberOfDataSheets, columnTimestamp, columnDuration, columnXO2, columnOUR, rowHeader, rowDataStart)
    If (numberOfDataSheets > 1) Then
        For i = 2 To numberOfDataSheets
            Sheets("Data" & i).Select
            If Range(columnTimestamp & rowDataStart).Value <> "" Then
                'Range("A" & rowDataStart & ":AY" & Range("B" & rowDataStart).End(xlDown).Row).Copy
                Range(columnTimestamp & rowDataStart & ":" & columnOUR & Range(columnDuration & rowDataStart).End(xlDown).Row).Copy
                Sheets("Data1").Select
                Columns(columnTimestamp & ":" & columnTimestamp).Select
                Selection.End(xlDown).Offset(1, 0).Select
                ActiveSheet.Paste
            End If
        Next
    End If
    
    'Round each timepoint, makes time synchronization easier
    lastRow = Application.WorksheetFunction.CountA(Columns(1))
    timeArray = Worksheets("Data1").Range(columnTimestamp & rowDataStart & ":" & columnTimestamp & lastRow).Value

    For j = 1 To UBound(timeArray, 1)
        timeArray(j, 1) = "=MROUND(""" & timeArray(j, 1) & """, ""0:01"")"
    Next
  
    Worksheets("Data1").Range(columnTimestamp & rowDataStart & ":" & columnTimestamp & lastRow).Value = timeArray

    'Remove number "1" from headers
    Worksheets("Data1").Range(columnTimestamp & rowHeader & ":" & columnXO2 & rowHeader).Replace What:="1", Replacement:=""
End Sub

'Sub will import OUR data
Private Sub importOURData(dasgipRawDataFileName, columnOURTime, columnCER, columnTimestamp, rowHeader, rowDataStart)
    Dim rawDataWorkbook As Workbook, targetWorkbook As Workbook
    Dim rawDataSheet As Worksheet, targetSheet As Worksheet
    Dim yy As String, mm As String, dd As String, ddOriginal As String, filter As String, fileFound As String, vesselID As String, lastColumnOUR As String
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
    lastColumnOUR = "M"

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
                lastRowTarget = targetSheet.Range(columnTimestamp & Rows.Count).End(xlUp).Row
                lastRowRaw = rawDataSheet.Range(columnTimestamp & Rows.Count).End(xlUp).Row

                If Not hasExistingData Then
                    'Copy data from OUR raw files to JMP Macro, no appending logic needed for the first day
                    targetSheet.Range(columnTimestamp & rowHeader, lastColumnOUR & lastRowRaw).Value = rawDataSheet.Range(columnTimestamp & rowHeader, lastColumnOUR & lastRowRaw).Value
                    hasExistingData = True
                Else
                    'Append to bottom if not the first day
                    lastRowTarget = lastRowTarget + 1
                    targetSheet.Range(columnTimestamp & lastRowTarget, lastColumnOUR & lastRowTarget + lastRowRaw - 2).Value = rawDataSheet.Range(columnTimestamp & rowDataStart, _
                        lastColumnOUR & lastRowRaw).Value
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
                        'mm = IIf(mm = "12", "01", CStr("0" & CInt(mm + 1)))
                        mm = IIf(mm = "12", "01", CStr(CInt(mm + 1)))
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
                      timeArrayOUR = .Range(columnTimestamp & rowDataStart & ":" & columnTimestamp & lastRowOUR).Value
                      
                      For j = 1 To UBound(timeArrayOUR, 1)
                          timeArrayOUR(j, 1) = "=MROUND(""" & timeArrayOUR(j, 1) & """, ""0:05"")"
                      Next
                    
                      .Range(columnTimestamp & rowDataStart & ":" & columnTimestamp & lastRowOUR).Value = timeArrayOUR
                    End With
                
                End If

                rawDataWorkbook.Close SaveChanges:=False

            Loop While hasAnotherDataFile = True
            
            'Determine data output intervals then space them out by this amount
            Call insertBlankRows(targetSheet, lastColumnOUR, lastRowOUR, columnTimestamp, rowDataStart)
        
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
                        
                        Debug.Print "DG Vessel: " & vesselID
                        
                        'Copy over matching OUR data to the corresponding row in the DG data
                        targetWorkbook.Worksheets("OUR" & vesselID).Range(columnTimestamp & matchedRowNumberOUR & ":" & lastColumnOUR & lastRowOUR).Copy _
                            Worksheets("Data" & i).Range(columnOURTime & matchedRowNumberDG & ":" & columnCER & (matchedRowNumberDG + lastRowOUR))
                            
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
        columnO2 As String, columnFPV As String, columnDG As String, columnStrainIDInput As String, columnCustomHeader1Input As String, _
        columnCustomHeader2Input As String
    Dim rawDataSheet As Worksheet, targetSheet As Worksheet
    Dim rawDataWorkbook As Workbook, targetWorkbook As Workbook
    Dim numberOfSpikes As Integer, lastRow As Integer, answer As Integer, totalRowCount As Integer, rowHeader As Integer, rowDataStart As Integer
    Dim importOUR As Boolean, isArgonLow As Boolean
    
    isArgonLow = False
    columnTimestamp = "A"
    columnDuration = "B"
    columnDG = "C"
    columnInoculationTime = "D"
    columnStrainIDInput = "E"
    columnCustomHeader1Input = "F"
    columnCustomHeader2Input = "G"
    columnFPV = "H"
    columnAlerts = "Q"
    columnTPV = "AC"
    columnVPV = "AE"
    columnVAPV = "AF"
    columnXCO2 = "AH"
    columnXO2 = "AI"
    columnStrainID = "AJ"
    columnCustomHeader1 = "AK"
    columnCustomHeader2 = "AL"
    columnNumberOfSpikes = "AM"
    columnBLI = "AN"
    columnOURTime = "AO"
    columnN2 = "AP"
    columnO2 = "AQ"
    columnAr = "AR"
    columnCO2 = "AS"
    columnOUR = "BB"
    columnCER = "BC"
    columnRQ = "BD"
    rowHeader = 1
    rowDataStart = 2
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
    Call removePreinoculationData(numberOfDataSheets, DG_Unit, rawDataWorkbook, targetWorkbook, columnDuration, columnDG, columnStrainIDInput, columnStrainID, _
        columnCustomHeader1Input, columnCustomHeader2Input, columnCustomHeader1, columnCustomHeader2, columnFPV, rowHeader, rowDataStart)
   
    'Copy data from DG raw files to JMP Macro
    For i = 1 To numberOfDataSheets
        Set targetSheet = targetWorkbook.Worksheets("Data" & i)
        Set rawDataSheet = rawDataWorkbook.Worksheets("Data" & i)
       
        'Identify last row in order to extract the correct range
        lastRow = Application.WorksheetFunction.CountA(rawDataSheet.Range(columnTimestamp & ":" & columnTimestamp))
       
        'DG3 and DG5 raw data export contain 6 additional columns than DG4, remove these columns so all DG units are formatted the same way
        'Update 02DEC2021 - This may  not be needed anymore
        If Application.WorksheetFunction.CountA(rawDataSheet.Range(columnBLI & ":" & columnBLI)) <> 0 Then
            rawDataSheet.Range("J:J,P:P,R:R,T:T,AL:AL,AN:AN").Delete
        End If
       
        With targetSheet
            'Import data
            .Range(columnTimestamp & rowDataStart, columnCustomHeader2 & lastRow).Value = rawDataSheet.Range(columnTimestamp & rowDataStart, columnCustomHeader2 & lastRow).Value
            
            If .Range(columnTimestamp & rowDataStart).Value <> "" Then
                'Round each timepoint, makes time synchronization easier
                timeArray = .Range(columnTimestamp & rowDataStart & ":" & columnCustomHeader2 & lastRow).Value
            
                For j = 1 To UBound(timeArray, 1)
                    timeArray(j, 1) = "=MROUND(""" & timeArray(j, 1) & """, ""0:01"")"
                Next
              
                .Range(columnTimestamp & rowDataStart & ":" & columnTimestamp & lastRow).Value = timeArray
               
                'Get number of DO spikes
                numberOfSpikes = countDOSpikes(lastRow, rawDataSheet, columnVAPV, rowDataStart)
                targetSheet.Range(columnNumberOfSpikes & rowDataStart & ":" & columnNumberOfSpikes & lastRow).Value = numberOfSpikes
            End If
        End With
    Next
   
    'Close raw data file
    rawDataWorkbook.Close SaveChanges:=False
    
    '#####################
    'TODO: Add BLI data
    '#####################
   
    'Import OUR data if selected
    answer = MsgBox("Would you like to import OUR data? (This may take a few minutes.)", vbYesNo)
    importOUR = IIf(answer = 6, True, False)
    
    If importOUR Then
        Call importOURData(rawDataFileName, columnOURTime, columnCER, columnTimestamp, rowHeader, rowDataStart)
    End If

    'Append all DG raw data to bottom of first sheet
    Call consolidateData(numberOfDataSheets, columnTimestamp, columnDuration, columnXO2, columnOUR, rowHeader, rowDataStart)
   
    'Convert Duration to array, perform "[h]:mm:ss" conversion, insert back into spreadsheet
    lastRow = Application.WorksheetFunction.CountA(Columns(1))
    timeArray = Worksheets("Data1").Range(columnDuration & rowDataStart & ":" & columnDuration & lastRow).Value
 
    For i = 1 To UBound(timeArray, 1)
        timeArray(i, 1) = "=TEXT(""" & timeArray(i, 1) & """, ""[h]:mm:ss"")"
    Next
   
    With Worksheets("Data1")
        .Range(columnDuration & rowDataStart & ":" & columnDuration & lastRow).Value = timeArray
    End With
   
    'Convert InoculationTime to array, perform text conversion, insert back into spreadsheet
    timeArray = Worksheets("Data1").Range(columnInoculationTime & rowDataStart & ":" & columnInoculationTime & lastRow).Value
 
    For i = 1 To UBound(timeArray, 1)
        timeArray(i, 1) = "=TEXT(""" & timeArray(i, 1) & """, ""[h]:mm:ss"")"
    Next
   
    With Worksheets("Data1")
       .Range(columnInoculationTime & rowDataStart & ":" & columnInoculationTime & lastRow).Value = timeArray
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
        Call fillColumnFromAbove(lastRow, columnFPV, columnTPV, columnVPV, columnXCO2, columnXO2)
        
        'Fix data types that tend to revert
        Range(columnCustomHeader1 & rowDataStart & ":" & columnCustomHeader2 & lastRow).NumberFormat = "General"
        Range(columnN2 & rowDataStart & ":" & columnN2 & lastRow).NumberFormat = "General"
        
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
        With Worksheets("Cover Sheet").Range(columnAlerts & "4")
            .Value = IIf(isArgonLow, "Low Argon levels detected", "Argon levels normal")
            .Interior.ColorIndex = IIf(isArgonLow, 3, 0)
        End With
        
        
        'Save workbook so the OUR/CER/RQ data shows up
        ActiveWorkbook.Save
    End If
End Sub
 
'Main macro container
Sub Run_JMP_Macro()
    Dim columnAlerts As String, columnCustom1 As String, columnCustom2 As String, columnLast As String

    columnAlerts = "Q"
    columnCustomHeader1 = "AK"
    columnCustomHeader2 = "AL"
    columnLast = "BD"
    
    'Clear pre-existing data except Cover Sheet
    For i = 2 To 17
        Worksheets(i).Range("A2:" & columnLast & Rows.Count).ClearContents
    Next

    'Clear colors and custom headers
    With Worksheets("Data1")
        .Range("A2:" & columnLast & Rows.Count).Interior.ColorIndex = 0
        .Range(columnCustomHeader1 & "1:" & columnCustomHeader2 & "1").ClearContents
        
        'Initialize headers
        .Range("A1:" & columnLast & "1").Value = Array( _
            "Timestamp", "Duration", "DG", "InoculationTime", _
            "DO.Out [%]", "DO.PV [%DO]", "DO.SP [%DO]", _
            "F.PV [sL/h]", "F.SP [sL/h]", "FA.PV [mL/h]", _
            "FAir.PV [sL/h]", "FAir.SP [sL/h]", _
            "FB.PV [mL/h]", "FB.SP [mL/h]", _
            "FCO2.PV [sL/h]", "FN2.PV [sL/h]", "FO2.PV [sL/h]", _
            "N.PV [rpm]", "N.SP [rpm]", "N.TStirPV [mNm]", _
            "OfflineA.OfflineA []", "OfflineB.OfflineB []", "OfflineC.OfflineC []", "OfflineD.OfflineD []", _
            "pH.Out [%]", "pH.PV [pH]", "pH.SP [pH]", _
            "T.Out [%]", "T.PV [°C]", "T.SP [°C]", _
            "V.VPV [mL]", "VA.PV [mL]", "VB.PV [mL]", _
            "XCO2.PV [%]", "XO2.PV [%]", _
            "Strain ID", "Custom 1", "Custom 2", _
            "Number of Spikes", "BLI", "Time", _
            "N2", "O2", "Ar", "CO2", "RMS Flow", "CDC", "OXQ", "RQ", "mass28", "mass32", "mass40", "mass44", "OUR", "CER", "RQ")
    End With
    
    'Clear alerts
    Worksheets("Cover Sheet").Range(columnAlerts & "4:" & columnAlerts & "27").ClearContents
    
    'Import DG raw data file
    Call importRawData
End Sub
