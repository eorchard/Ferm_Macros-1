Private Sub Workbook_Open()
 
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.Calculation = xlManual
  
'##############################################################################################################
'### DASGIP DATA IMPORT #######################################################################################
'##############################################################################################################
  
    
Dim ImportFilePath As Variant
Dim ImportFileName As Variant
Dim book As Workbook
Dim importsheet As Worksheet
Dim exportsheet As Worksheet
Dim start As Variant
 
'### Sets up worksheet open ###################################################################################
 
start = MsgBox("You will be prompted to select the DASgip data you'd like to graph. Please select the relevant control file that is exported by the DASgip software. This macro works with data exported by DASware 5.0 and will not work with the older DASgip software.", vbOKCancel)
    If start = vbCancel Then Exit Sub
   
    ImportFilePath = Application.GetOpenFilename(, , "Which DASgip file would you like to import and graph?")
    ImportFileName = Right(ImportFilePath, Len(ImportFilePath) - InStrRev(ImportFilePath, "\"))
    If ImportFilePath = False Then
        Exit Sub
    Else: Workbooks.Open (ImportFilePath)
    End If
 
'### Copies cell values from import workbook to export work book (this workbook)and marks with "."############
 
        For Each importsheet In Workbooks(ImportFileName).Worksheets
            For Each exportsheet In ThisWorkbook.Worksheets
                If exportsheet.Name = importsheet.Name Then
                    Workbooks(ImportFileName).Worksheets(importsheet.Name).Cells.Copy Destination:=ThisWorkbook.Worksheets(exportsheet.Name).Cells
                    ThisWorkbook.Worksheets(exportsheet.Name).Name = exportsheet.Name & "." '### Marks Copied cells with "."
                Else 'do nothing
                End If
                Next exportsheet
        Next importsheet
               
Workbooks(ImportFileName).Close
 
'### Deletes all sheets without "." in them and then removes "." from each sheet title#########################
            ThisWorkbook.Worksheets("Data Handling Tools").Name = "Data Handling Tools" & "."
           
            For Each exportsheet In ThisWorkbook.Worksheets
                If InStr(ThisWorkbook.Worksheets(exportsheet.Name).Name, ".") = 0 Then
                    ThisWorkbook.Worksheets(exportsheet.Name).Delete
                Else: 'do nothing
                End If
            Next exportsheet
           
            For Each exportsheet In ThisWorkbook.Worksheets
                If InStr(exportsheet.Name, ".") = 0 Then
                    'do nothing
                Else: ThisWorkbook.Worksheets(exportsheet.Name).Name = Left(Worksheets(exportsheet.Name).Name, Len(Worksheets(exportsheet.Name).Name) - 1)
                End If
            Next exportsheet
 
'##############################################################################################################
'### OUR IMPORT AND ROW FILL ##################################################################################
'##############################################################################################################
 
 
'### Determines if offgas import should happen and then starts process########################################
Dim importoffgas As Integer
importoffgas = MsgBox("Would you like to import bluesens offgas data and calculate OUR?", vbYesNo)
    If importoffgas = vbNo Then
        GoTo CloseMacro
    End If
 
MsgBox ("Please select the bluesens offgas data file (.xlxs)")
    ImportFilePath = Application.GetOpenFilename(, , "Which bluesens offgas data would you like to import?")
    ImportFileName = Right(ImportFilePath, Len(ImportFilePath) - InStrRev(ImportFilePath, "\"))
    If ImportFilePath = False Then '###Major If###
        'ThisWorkbook.Close
        Exit Sub
    Else: 'do nothing
    End If
MsgBox ("This will take a decent amount of computing power, please don't touch the screen or exit the program.")
   
'### Sets up export sheet cycling #############################################################################
 
Dim i As Integer
Dim n As Integer
Dim startrow As Integer
Dim totalrowcount As Integer
Dim totalinocpoints As Integer
Dim count As Integer
    count = 0
Dim col As Integer
 
For Each exportsheet In ThisWorkbook.Worksheets
If InStr(exportsheet.Name, "Data") = 0 Then '### If A ###
    Else:
        If IsEmpty(ThisWorkbook.Worksheets(exportsheet.Name).Cells(1, 1).Value) = True Then '###If Statement B#
        Else:
        count = count + 1
 
With ThisWorkbook.Worksheets(exportsheet.Name) '### With A ###
                   
                    
'### Defines total rows and start of inoculation time #########################################################
    totalrowcount = .Columns(1).End(xlDown).Row
        For i = 2 To totalrowcount
            If IsEmpty(.Cells(i, 3).Value) = False Then
                startrow = i
                totalinocpoints = totalrowcount - startrow
                i = totalrowcount
            Else 'do nothing
            End If
        Next i
                    
'###Adds offgas column and OUR column w/formula################################################################
    .Cells(1, 35).Value = "Offgas %O2"
    .Cells(1, 36).Value = "OUR (mmoL/L/hr)"
    For i = startrow To totalrowcount
        .Cells(i, 36).Value = "=(((G" & i & "/60)*(AH" & i & "-AI" & i & ")*600)/((AD" & i & "/1000)*0.082057*(AB" & i & "+273.15)))"
    Next i
                   
'### Fills Empty Rows (select row to save computing power) ####################################################
    For i = startrow - 10 To totalrowcount
                       
    '##Flow Rate = G = 7##
    If .Cells(i, 7).Value = 0 Then
    .Cells(i, 7).Value = .Cells(i - 1, 7).Value
    End If
                       
    '##O2 Supplement = AH = 34##
    If .Cells(i, 34).Value = 0 Then
    .Cells(i, 34).Value = .Cells(i - 1, 34).Value
    End If
                       
    '##Offgas % O2 = AI = 35## Currently disabled because not needed
    If .Cells(i, 35).Value = 0 Then
    .Cells(i, 35).Value = .Cells(i - 1, 35).Value
    End If
                       
    '##Tank Volume = AD = 30##
    If .Cells(i, 30).Value = 0 Then
    .Cells(i, 30).Value = .Cells(i - 1, 30).Value
    End If
                        
    '##Temperature = AD = 28##
    If .Cells(i, 28).Value = 0 Then
    .Cells(i, 28).Value = .Cells(i - 1, 28).Value
    End If
   
    Next i
   
End With
   
'### Imports offgas data and pastes in the workbook ###########################################################
col = count * 2 + 2
 
Workbooks.Open (ImportFilePath) '### Checks to see if values are 30s apart, implement later if problem. If (workbooks(importfilename).Worksheets(1).Cells(4,3).value - workbooks(importfilename).Worksheets(1).Cells(3,3).value) = 0.0083333
Workbooks(ImportFileName).Worksheets(1).Range(Cells(3, col), Cells(3 + totalinocpoints, col)).Copy _
Destination:=ThisWorkbook.Worksheets(exportsheet.Name).Cells(startrow, 35)
Workbooks(ImportFileName).Close
 
    End If '### If B###
    End If '### If A###
   
Next exportsheet
 
 
'### Row Fill for all rows. Works but takes a lot of computing power. Currently Disabled  #####################
'##############################################################################################################
 
'    For Each exportsheet In ThisWorkbook.Worksheets
'        If InStr(exportsheet.Name, "Data") = 0 Then 'do nothing
'        Else:
'            If IsEmpty(ThisWorkbook.Worksheets(exportsheet.Name).Cells(1, 1).Value) = True Then ' do nothing
'            Else:
'                With ThisWorkbook.Worksheets(exportsheet.Name)
'                    totalrowcount = .Columns(1).End(xlDown).Row
'                        For n = 1 To 34 '##34 is last column from DASGIP import
'                            For i = 10 To totalrowcount
'                                If .Cells(i, n).Value = 0 Then
'                                    .Cells(i, n).Value = .Cells(i - 1, n).Value
'                                End If
'                            Next i
'                        Next n
'                End With
'            End If
'        End If
'    Next exportsheet
'################################################################################################################
 
CloseMacro:
   
MsgBox ("Charts are now displayed. To view data or event history unhide the relevant worksheet tab.")
 
 
Application.Calculation = xlAutomatic
Application.ScreenUpdating = True
Application.DisplayAlerts = True
 
 
End Sub
