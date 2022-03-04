'Makes string comparisons case insensitive
Option Compare Text

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
    
    'Attempt to improve performance since this subroutine is time-consuming
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
        
        'The multiple will determine how many rows to add
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
 
'Sub adds vessel column to raw data
Private Sub addVesselColumn(vessel, vesselNumber, rawDataWorkbook, targetWorkbook, columnDuration, columnFermenter, columnStrainIDInput, columnStrainID, _
        columnCustomHeader1Input, columnCustomHeader2Input, columnCustomHeader1, columnCustomHeader2, rowFirst, rowDataStart)
        
    Dim strainID As String, customHeader1 As String, customHeader2 As String, customHeaderData1 As String, customHeaderData2 As String
    Dim rowHeader As Integer, rowOffset As Integer
    
    rowHeader = 3
    rowOffset = -1

    With targetWorkbook.Worksheets("Cover Sheet")
        customHeader1 = .Range(columnCustomHeader1Input & rowHeader).Value
        customHeader2 = .Range(columnCustomHeader2Input & rowHeader).Value
        
        If vessel = "DG3_u" Then
            rowOffset = 0
        ElseIf vessel = "DG4_u" Then
            rowOffset = 8
        ElseIf vessel = "DG5_u" Then
            rowOffset = 16
        ElseIf vessel = "Appalachian" Then
            rowOffset = 25
        ElseIf vessel = "Brooks" Then
            rowOffset = 26
        ElseIf vessel = "Cascades" Then
            rowOffset = 27
        ElseIf vessel = "Dolomites" Then
            rowOffset = 28
        ElseIf vessel = "Elk" Then
            rowOffset = 29
        ElseIf vessel = "Himalayas" Then
            rowOffset = 30
        End If
        
        strainID = .Range(columnStrainIDInput & (rowHeader + rowOffset) + vesselNumber).Value
        customHeaderData1 = .Range(columnCustomHeader1Input & (rowHeader + rowOffset) + vesselNumber).Value
        customHeaderData2 = .Range(columnCustomHeader2Input & (rowHeader + rowOffset) + vesselNumber).Value
        
        If rowOffset = -1 Then
            strainID = "N/A"
            customHeaderData1 = "N/A"
            customHeaderData2 = "N/A"
        End If
    End With
    
    ActiveWindow.ScrollRow = 1
    
    'Name custom headers
    With targetWorkbook.Worksheets("Data1")
        .Range(columnCustomHeader1 & rowFirst).Value = customHeader1
        .Range(columnCustomHeader2 & rowFirst).Value = customHeader2
    End With
    
    'Fill columns
    With Range(columnFermenter & rowDataStart & ":" & columnFermenter & Range(columnDuration & rowDataStart).End(xlDown).Row)
        .Value = IIf(vesselNumber = 0, vessel, vessel & vesselNumber)
    End With
            
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
Private Sub removePreinoculationData(numberOfDataSheets, fermenterName, rawDataWorkbook, targetWorkbook, columnDuration, columnFermenter, columnStrainIDInput, columnStrainID, _
        columnCustomHeader1Input, columnCustomHeader2Input, columnCustomHeader1, columnCustomHeader2, columnFPV, rowHeader, rowDataStart, is30L)

    If is30L Then
        targetWorkbook.Sheets("Data1").Select
        Columns("W:W").Select
        Selection.End(xlDown).Select
        Range("A2:A" & ActiveCell.Row - 1).EntireRow.Delete
        Call addVesselColumn(fermenterName, 0, rawDataWorkbook, targetWorkbook, columnDuration, columnFermenter, columnStrainIDInput, columnStrainID, _
            columnCustomHeader1Input, columnCustomHeader2Input, columnCustomHeader1, columnCustomHeader2, rowHeader, rowDataStart)
    Else
    
        For i = 1 To numberOfDataSheets
            Sheets("Data" & i).Select
            'Need this condition to check if Data Sheet # matches actual vessel since the vessels aren't always run in sequence
            If Range(columnFPV & rowHeader).Value Like "*" & i & "*" Then
                Columns("C:C").Select
                Selection.End(xlDown).Select
                Range("A2:A" & ActiveCell.Row - 1).EntireRow.Delete
                Call addVesselColumn(fermenterName, i, rawDataWorkbook, targetWorkbook, columnDuration, columnFermenter, columnStrainIDInput, columnStrainID, _
                    columnCustomHeader1Input, columnCustomHeader2Input, columnCustomHeader1, columnCustomHeader2, rowHeader, rowDataStart)
            
            'If vessel doesn't match up with Data Sheet #, compare the other numbers
            Else
                For j = 1 To 8
                    If Range(columnFPV & rowHeader).Value Like "*" & j & "*" Then
                        Columns("C:C").Select
                        Selection.End(xlDown).Select
                        Range("A2:A" & ActiveCell.Row - 1).EntireRow.Delete
                        Call addVesselColumn(fermenterName, j, rawDataWorkbook, targetWorkbook, columnDuration, columnFermenter, columnStrainIDInput, columnStrainID, _
                            columnCustomHeader1Input, columnCustomHeader2Input, columnCustomHeader1, columnCustomHeader2, rowHeader, rowDataStart)
                    End If
                Next
            End If
        Next
    End If
End Sub
 
'Sub will consolidate data onto one sheet
Private Sub consolidateData(numberOfDataSheets, columnTimestamp, columnDuration, columnXO2, columnOUR, rowHeader, rowDataStart)
    If (numberOfDataSheets > 1) Then
        For i = 2 To numberOfDataSheets
            Sheets("Data" & i).Select
            If Range(columnTimestamp & rowDataStart).Value <> "" Then
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
Private Sub importOURData(rawDataFileName, columnOURTime, columnCER, columnTimestamp, rowHeader, rowDataStart, is30L, Optional ByVal fermenterName)
    Dim rawDataWorkbook As Workbook, targetWorkbook As Workbook
    Dim rawDataSheet As Worksheet, targetSheet As Worksheet
    Dim yy As String, mm As String, dd As String, ddOriginal As String, filter As String, fileFound As String, vesselID As String, lastColumnOUR As String
    Dim numberOfDaysPerMonthArray As Variant
    Dim lastRowTarget As Integer, lastRowRaw As Integer, lastRowOUR As Integer, matchedRowNumberOUR As Integer, matchedRowNumberFermenter As Integer, _
        dasgipID As Integer
    Dim hasAnotherDataFile As Boolean, hasExistingData As Boolean, exitFor As Boolean
    Dim datePattern As Object, datePatternRegExp As Object, dasgipIDPattern As Object, vesselIDRegExp As Object
    Dim firstTimeToMatch As Double
    
    numberOfDaysPerMonthArray = Array(31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31)
    Set datePatternRegExp = New RegExp
    Set vesselIDRegExp = New RegExp
    fileNotFound = True
    lastColumnOUR = "M"

    'Get OUR raw data
    filter = "Text files (*.xlsx),*.xlsx"
    Set targetWorkbook = Application.ThisWorkbook

    'Parse OUR date to query based off of date from raw data filename
    datePatternRegExp.Pattern = "\d{6}"
    Set datePattern = datePatternRegExp.Execute(rawDataFileName)
    dd = Right(datePattern(0), 2)
    mm = Mid(datePattern(0), 3, 2)
    yy = Left(datePattern(0), 2)
    ddOriginal = dd
    mmOriginal = mm
        
    If Not is30L Then
        'Find the DG file using pattern matching, file must belong in a folder called 'raw data'
        'Pattern must include "raw data\" to exclude false matches from the batch ID
        'Example filename: 220202DG5 raw process data
        vesselIDRegExp.Pattern = "raw data\\.*DG\d"
        Set dasgipIDPattern = vesselIDRegExp.Execute(rawDataFileName)
        dasgipID = Right(dasgipIDPattern(0), 1)
    End If
    
    'Copy paste data from each OUR file into sheet tab (OUR1, OUR2, etc..)
    For i = 1 To 8
        hasExistingData = False

        'Parse cell C2 to get the correct DG unit; can't rely on worksheet name. Use vessel number to open raw data file
        If Worksheets("Data" & i).Cells(2, 3) <> "" And Not is30L Then
            vesselID = Right(Worksheets("Data" & i).Cells(2, 3), 1)
            fileFound = Dir("S:\Projects\Fermentation\Ferm&StrainDevelopment\OUR Data\20" & yy & "\" & dasgipID & "-" & vesselID & "\analysis\" & mm & dd & "*.csv")
            Set targetSheet = targetWorkbook.Worksheets("OUR" & vesselID)
        ElseIf is30L Then
            vesselID = 1
            fileFound = Dir("S:\Projects\Fermentation\Ferm&StrainDevelopment\OUR Data\20" & yy & "\" & "30L - 3" & "\analysis\" & mm & dd & "*.csv")
            Set targetSheet = targetWorkbook.Worksheets("OUR1")
        End If
            
        'Run block if file exists
        If fileFound <> "" Then
            fileNotFound = False
                        
            'Collect OUR data for individual vessel until no more sequential data files exist
            Do
                If is30L Then
                    Set rawDataWorkbook = Application.Workbooks.Open("S:\Projects\Fermentation\Ferm&StrainDevelopment\OUR Data\20" & yy & "\" & _
                        "30L - 3" & "\analysis\" & mm & dd & "*.csv")
                Else
                    Set rawDataWorkbook = Application.Workbooks.Open("S:\Projects\Fermentation\Ferm&StrainDevelopment\OUR Data\20" & yy & "\" & dasgipID & _
                      "-" & vesselID & "\analysis\" & mm & dd & "*.csv")
                End If
                
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
                
                'Cycle dd when exceeds the max value
                If (CInt(dd) > numberOfDaysPerMonthArray(CInt(mm) - 1)) Then
                    dd = "01"
                    'Cycle mm
                    If (CInt(mm) < CInt("09")) Then
                        mm = CStr("0" & CInt(mm + 1))
                    ElseIf mm = "09" Then
                        mm = "10"
                    ElseIf (CInt(mm) > CInt("09")) Then
                        mm = CStr(CInt(mm + 1))
                    End If
                End If

                'Check if an OUR data file for the next day exists
                fileFound = IIf(is30L, Dir("S:\Projects\Fermentation\Ferm&StrainDevelopment\OUR Data\20" & yy & "\" & "30L - 3" _
                & "\analysis\" & mm & dd & "*.csv"), Dir("S:\Projects\Fermentation\Ferm&StrainDevelopment\OUR Data\20" & yy & "\" & dasgipID & _
                    "-" & vesselID & "\analysis\" & mm & dd & "*.csv"))
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
                For k = 3 To 1440
                    If targetWorkbook.Worksheets("Data" & i).Cells(j, 1).Value = targetWorkbook.Worksheets("OUR" & vesselID).Cells(k, 1).Value Then
                        'Matching timepoints
                        matchedRowNumberOUR = k
                        matchedRowNumberFermenter = j
                        
                        'Copy over matching OUR data to the corresponding row in the DG data
                        targetWorkbook.Worksheets("OUR" & vesselID).Range(columnTimestamp & matchedRowNumberOUR & ":" & lastColumnOUR & lastRowOUR).Copy _
                            Worksheets("Data" & i).Range(columnOURTime & matchedRowNumberFermenter & ":" & columnCER & (matchedRowNumberFermenter + lastRowOUR))
                            
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
            End
        End If
    Next
End Sub
 
'Function imports raw data file from DG units
Private Sub import2LRawData(columnAr As String, columnCER, columnCO2, columnCustom1, columnCustom2, columnCustomHeader1, columnCustomHeader2, _
        columnCustomHeader1Input, columnCustomHeader2Input, columnDuration, columnFermenter, columnFPV, columnInoculationTime, columnN2, columnNumberOfSpikes, columnO2, columnOUR, columnOURTime, columnRQ, columnStrainID, _
        columnStrainIDInput, columnTimestamp, columnTPV, columnVAPV, columnVPV, columnXCO2, columnXO2, rowHeader, rowDataStart, columnAlerts, is30L)
    Dim filter As String, DG_Unit As String
    Dim rawDataSheet As Worksheet, targetSheet As Worksheet
    Dim rawDataWorkbook As Workbook, targetWorkbook As Workbook
    Dim numberOfSpikes As Integer, lastRow As Integer, answer As Integer
    Dim importOUR As Boolean, isArgonLow As Boolean

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
        Else
            MsgBox "Filename not valid, filename needs to include: 'DG3' 'DG4' or 'DG5'"
            End
        End If
    End With
    
    'Remove six columns below, these were added when the DASGIP software was upgraded. Makes the macro backward-compatible
    For i = 1 To numberOfDataSheets
        Set rawDataSheet = rawDataWorkbook.Worksheets("Data" & i)
        If Application.WorksheetFunction.CountA(rawDataSheet.Range("AN:AN")) <> 0 Then
            rawDataSheet.Range("J:J,P:P,R:R,T:T,AL:AL,AN:AN").Delete
        End If
    Next
   
    'Remove timepoints before inoculation from raw data
    Call removePreinoculationData(numberOfDataSheets, DG_Unit, rawDataWorkbook, targetWorkbook, columnDuration, columnFermenter, columnStrainIDInput, columnStrainID, _
        columnCustomHeader1Input, columnCustomHeader2Input, columnCustomHeader1, columnCustomHeader2, columnFPV, rowHeader, rowDataStart, is30L)
   
    'Copy data from DG raw files to JMP Macro
    For i = 1 To numberOfDataSheets
        Set targetSheet = targetWorkbook.Worksheets("Data" & i)
        Set rawDataSheet = rawDataWorkbook.Worksheets("Data" & i)
       
        'Identify last row in order to extract the correct range
        lastRow = rawDataSheet.Range(columnTimestamp & Rows.Count).End(xlUp).Row
       
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
   
    'Import OUR data if selected
    answer = MsgBox("Would you like to import OUR data? (This may take a few minutes.)", vbYesNo)
    importOUR = IIf(answer = 6, True, False)
    
    If importOUR Then
        Call importOURData(rawDataFileName, columnOURTime, columnCER, columnTimestamp, rowHeader, rowDataStart, is30L)
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
        Range(columnCustomHeader1 & ":" & columnCustomHeader2).NumberFormat = "General"
        Range(columnN2 & ":" & columnN2).NumberFormat = "General"
        
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

'Function imports raw data file from Sartorius fermenters
Private Sub import30LRawData(columnAr, columnCER, columnCO2, columnCustom1, columnCustom2, columnCustomHeader1, columnCustomHeader2, columnCustomHeader1Input, columnCustomHeader2Input, _
            columnDuration, columnFermenter, columnFPV, columnInoculationTime, columnInoculationTimeInput, columnN2, columnNumberOfSpikes, columnO2PV, columnOUR, columnOURTime, columnRQ, columnStirr, columnStrainID, _
            columnStrainIDInput, columnTimestamp, columnTPV, columnVAPV, columnVPV, columnXCO2, columnXO2, columnPHSP, columnPHPV, columnO2SP, columnPressure, _
            columnVBPV, columnBaseTValue, columnTSP, columnAFOMT, rowHeader, rowDataStart, columnAlerts, is30L)
    Dim filter As String, fermenterName As String, rowInoculationTime As String
    Dim rawDataSheet As Worksheet, targetSheet As Worksheet
    Dim rawDataWorkbook As Workbook, targetWorkbook As Workbook
    Dim numberOfSpikes As Integer, lastRow As Integer, answer As Integer
    Dim importOUR As Boolean, isArgonLow As Boolean
    Dim inoculationTime As Variant

    filter = "Text files (*.xls),*.xls"
    MsgBox "Please select the 30L/100L raw data file", vbOKOnly
   
    'JMP Macro workbook is the target
    Set targetWorkbook = Application.ThisWorkbook
    Set targetSheet = targetWorkbook.Worksheets("Data1")
   
    'Get raw data workbook
    rawDataFileName = Application.GetOpenFilename(filter, , Caption)
    
    If rawDataFileName = False Then
        Exit Sub
    Else
        Set rawDataWorkbook = Application.Workbooks.Open(rawDataFileName)
        Set rawDataSheet = rawDataWorkbook.Worksheets(1)
    End If
   
    With rawDataWorkbook
    'Identify which fermenter the raw data is coming from, currently relies on filename
        If .Name Like "*" & "Appalachian" & "*" Then
            fermenterName = "Appalachian"
            rowInoculationTime = 28
        ElseIf .Name Like "*" & "Brooks" & "*" Then
            fermenterName = "Brooks"
            rowInoculationTime = 29
        ElseIf .Name Like "*" & "Cascades" & "*" Then
            fermenterName = "Cascades"
            rowInoculationTime = 30
        ElseIf .Name Like "*" & "Dolomites" & "*" Then
            fermenterName = "Dolomites"
            rowInoculationTime = 31
        ElseIf .Name Like "*" & "Elk" & "*" Then
            fermenterName = "Elk"
            rowInoculationTime = 32
        ElseIf .Name Like "*" & "Himalayas" & "*" Then
            fermenterName = "Himalayas"
            rowInoculationTime = 33
        Else
            MsgBox "Filename not valid, filename needs to include: 'Appalachian' 'Brooks' 'Cascades' 'Dolomites' 'Elk' or 'Himalayas'"
            End
        End If
    End With
   
     'Identify last row in order to extract the correct range
        lastRow = rawDataSheet.Range(columnTimestamp & Rows.Count).End(xlUp).Row

    With targetSheet
        'Import data
        .Range(columnTimestamp & rowDataStart, columnCustomHeader2 & lastRow).Value = rawDataSheet.Range(columnTimestamp & rowDataStart, columnCustomHeader2 & lastRow).Value
        
        'When no more data files exist, round all the OUR timepoints to nearest fifth minute. Makes time synchronization easier
        timeArray = .Range(columnTimestamp & rowDataStart & ":" & columnTimestamp & lastRow).Value
        
        For j = 1 To UBound(timeArray, 1)
            timeArray(j, 1) = "=MROUND(""" & timeArray(j, 1) & """, ""0:01"")"
        Next

        .Range(columnTimestamp & rowDataStart & ":" & columnTimestamp & lastRow).Value = timeArray
    End With

    'Match timepoints on raw data when the fermenter was inoculated
    For j = 2 To lastRow
        If CStr(targetWorkbook.Worksheets(1).Cells(rowInoculationTime, columnInoculationTimeInput).Value) = CStr(targetWorkbook.Worksheets("Data1").Cells(j, 1).Value) Then
            'Matched timepoints, fill down inoculation time starting from 0:00 in 1-minute increments
            With targetWorkbook.Worksheets("Data1")
                .Range(columnInoculationTime & j).Value = "0:00"
                For k = j To lastRow - 1
                    .Range(columnInoculationTime & k + 1).Value = DateAdd("n", 1, .Range(columnInoculationTime & k).Value)
                Next
            End With
            Exit For
        End If
    Next
    
    'Check if raw column headers match what is expected, otherwise throw an error
    If rawDataSheet.Range(columnFPV & rowHeader).Value <> "AIR_SP_Value" Or _
        rawDataSheet.Range(columnXO2 & rowHeader).Value <> "O2_SP_Value" Or _
        rawDataSheet.Range(columnPHSP & rowHeader).Value <> "pH_Setpoint" Or _
        rawDataSheet.Range(columnPHPV & rowHeader).Value <> "pH_Value" Or _
        rawDataSheet.Range(columnO2SP & rowHeader).Value <> "pO2_Setpoint" Or _
        rawDataSheet.Range(columnO2PV & rowHeader).Value <> "pO2_Value" Or _
        rawDataSheet.Range(columnPressure & rowHeader).Value <> "PRESS_Value" Or _
        rawDataSheet.Range(columnStirr & rowHeader).Value <> "STIRR_Value" Or _
        rawDataSheet.Range(columnBaseTValue & rowHeader).Value <> "BASET_Value" Or _
        rawDataSheet.Range(columnTPV & rowHeader).Value <> "TEMP_Value" Or _
        rawDataSheet.Range(columnTSP & rowHeader).Value <> "TEMP_Setpoint" Or _
        rawDataSheet.Range(columnAFOMT & rowHeader).Value <> "AFOMT_Value" Then
        MsgBox "30L/100L column headers don't match the expected format, please see the expected order of columns"
        End
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
     
        'Close raw data file
        rawDataWorkbook.Close SaveChanges:=False
         
        'Remove timepoints before inoculation from raw data
        Call removePreinoculationData(1, fermenterName, rawDataWorkbook, targetWorkbook, columnDuration, columnFermenter, columnStrainIDInput, columnStrainID, _
            columnCustomHeader1Input, columnCustomHeader2Input, columnCustomHeader1, columnCustomHeader2, columnFPV, rowHeader, rowDataStart, is30L)
        
    End With
   
    'Import OUR data if selected
    'answer = MsgBox("Would you like to import OUR data? (This may take a few minutes.)", vbYesNo)
    'importOUR = IIf(answer = 6, True, False)
    importOUR = False
    
    If importOUR Then
        Call importOURData(rawDataFileName, columnOURTime, columnCER, columnTimestamp, rowHeader, rowDataStart, is30L, fermenterName)
    End If
   
    'Convert Age to array, perform "General" conversion, insert back into spreadsheet
    timeArray = targetSheet.Range(columnDuration & rowDataStart & ":" & columnDuration & lastRow).Value

    For i = 1 To UBound(timeArray, 1)
        timeArray(i, 1) = "=TEXT(""" & timeArray(i, 1) & """, ""General"")"
    Next

    targetSheet.Range(columnDuration & rowDataStart & ":" & columnDuration & lastRow).Value = timeArray
    
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
            'Raise flag if Argon values are out of bounds
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

    Worksheets("Data1").Select
End Sub
 
'Main macro container for 2L
Sub Run_JMP_Macro(is30L)
    Dim columnAFOMT As String, columnAr As String, columnBaseTValue As String, columnCER As String, columnCO2 As String, columnCustom1 As String, columnCustom2 As String, _
        columnCustomHeader1 As String, columnCustomHeader2 As String, columnCustomHeader1Input As String, columnCustomHeader2Input As String, _
        columnDuration As String, columnFermenter As String, columnFPV As String, columnInoculationTime As String, columnO2 As String, columnO2SP As String, columnO2PV As String, columnN2 As String, columnNPV As String, _
        columnNumberOfSpikes As String, columnOUR As String, columnOURTime As String, columnPressure As String, columnRQ As String, columnStrainID As String, columnStrainIDInput As String, _
        columnTimestamp As String, columnTPV As String, columnVAPV As String, columnVPV As String, columnXCO2 As String, columnXO2 As String
    Dim rowHeader As Integer, rowDataStart As Integer
    Dim isArgonLow As Boolean
    
    isArgonLow = False
    columnTimestamp = "A"
    columnDuration = "B"
    columnFermenter = "C"
    columnInoculationTime = "D"
    columnStrainIDInput = "E"
    columnCustomHeader1Input = "F"
    columnCustomHeader2Input = "G"
    columnFPV = "H"
    columnAlerts = "Q"
    columnNPV = "R"
    columnPHPV = "Z"
    columnPHSP = "AA"
    columnTPV = "AC"
    columnTSV = "AD"
    columnVPV = "AE"
    columnVAPV = "AF"
    columnVBPV = "AG"
    columnXCO2 = "AH"
    columnXO2 = "AI"
    columnStrainID = "AJ"
    columnCustomHeader1 = "AK"
    columnCustomHeader2 = "AL"
    columnNumberOfSpikes = "AM"
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
    
    If is30L Then
        columnInoculationTimeInput = "C"
        columnFPV = "D"
        columnXO2 = "E"
        columnPHSP = "F"
        columnPHPV = "G"
        columnO2SP = "H"
        columnO2PV = "I"
        columnPressure = "J"
        columnStirr = "K"
        columnVAPV = "M"
        columnVBPV = "N"
        columnBaseTValue = "O"
        columnTPV = "P"
        columnTSP = "Q"
        columnAFOMT = "R"
        columnStrainID = "S"
        columnCustomHeader1 = "T"
        columnCustomHeader2 = "U"
        columnNumberOfSpikes = "V"
        columnInoculationTime = "W"
        columnOURTime = "X"
        columnN2 = "Y"
        columnO2 = "Z"
        columnAr = "AA"
        columnCO2 = "AB"
        columnOUR = "AK"
        columnCER = "AL"
        columnRQ = "AM"
    End If
    
    'Clear pre-existing data except Cover Sheet
    For i = 2 To 17
        Worksheets(i).Range("A1:BZ" & Rows.Count).ClearContents
    Next
    
    'Clear alerts
    Worksheets("Cover Sheet").Range(columnAlerts & "4:" & columnAlerts & "34").ClearContents
    
    'Clear colors and custom headers
    With Worksheets("Data1")
        .Range("A1:BZ" & Rows.Count).Interior.ColorIndex = 0
        .Range(columnCustomHeader1 & "1:" & columnCustomHeader2 & "1").ClearContents
    End With
    
    'Initialize headers
    If Not is30L Then
        'Initialize DG headers
        Worksheets("Data1").Range("A1:" & columnRQ & "1").Value = Array( _
            "Timestamp", "Duration", "Vessel", "InoculationTime", _
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
            "Number of Spikes", "", "Time", _
            "N2", "O2", "Ar", "CO2", "RMS Flow", "CDC", "OXQ", "RQ", "mass28", "mass32", "mass40", "mass44", "OUR", "CER", "RQ")
            
        'Import 2L raw data file
        Call import2LRawData(columnAr, columnCER, columnCO2, columnCustom1, columnCustom2, columnCustomHeader1, columnCustomHeader2, _
            columnCustomHeader1Input, columnCustomHeader2Input, columnDuration, columnFermenter, columnFPV, columnInoculationTime, columnN2, columnNumberOfSpikes, columnO2, columnOUR, columnOURTime, columnRQ, columnStrainID, _
            columnStrainIDInput, columnTimestamp, columnTPV, columnVAPV, columnVPV, columnXCO2, columnXO2, rowHeader, rowDataStart, columnAlerts, is30L)
            
    Else
        'Initialize 30L headers
        Worksheets("Data1").Range("A1:" & columnRQ & "1").Value = Array( _
            "PDatTime", "Age", "Vessel", "AIR_SP_Value", "O2_SP_Value", "pH_Setpoint", "pH_Value", "pO2_Setpoint", "pO2_Value", "PRESS_Value", "STIRR_Value", "SUBS_A_Value", _
            "SUBST_A_Value", "VWEIGHT_Value", "BASET_Value", "TEMP_Value", "TEMP_Setpoint", "AFOMT_Value", "Strain ID", "Custom 1", "Custom 2", _
            "Number of Spikes", "Inoculation Time", "Time", "N2", "O2", "Ar", "CO2", "RMS Flow", "CDC", "OXQ", "RQ", "mass28", "mass32", "mass40", "mass44", "OUR", "CER", "RQ")
            
        'Import 30L raw data file
        Call import30LRawData(columnAr, columnCER, columnCO2, columnCustom1, columnCustom2, columnCustomHeader1, columnCustomHeader2, columnCustomHeader1Input, columnCustomHeader2Input, _
            columnDuration, columnFermenter, columnFPV, columnInoculationTime, columnInoculationTimeInput, columnN2, columnNumberOfSpikes, columnO2PV, columnOUR, columnOURTime, columnRQ, columnStirr, columnStrainID, _
            columnStrainIDInput, columnTimestamp, columnTPV, columnVAPV, columnVPV, columnXCO2, columnXO2, columnPHSP, columnPHPV, columnO2SP, columnPressure, _
            columnVBPV, columnBaseTValue, columnTSP, columnAFOMT, rowHeader, rowDataStart, columnAlerts, is30L)
    End If
    
End Sub

'Main macro container for 30L/100L
Sub Run_JMP_Macro_30L()
    Dim is30L As Boolean
    
    is30L = True
    
    Call Run_JMP_Macro(is30L)
End Sub

'Main macro container for 2L
Sub Run_JMP_Macro_2L()
    Dim is30L As Boolean
    
    is30L = False
    
    Call Run_JMP_Macro(is30L)
End Sub

