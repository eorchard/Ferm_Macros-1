Private Sub parseBiolectorData()
    Dim rawDataWorkbook As Workbook, targetWorkbook As Workbook
    Dim rawDataSheet As Worksheet, targetSheet As Worksheet, mainSheet As Worksheet
    Dim patternFound As Range
    Dim filter As String, lastColumnChar As String, currentColumnChar As String, vesselID As String
    Dim headersArray As Variant, dataArray As Variant
    Dim timepointArray As Double
    Dim patternFoundRow As Integer, rawDataSectionFirstRow As Integer, rawDataSectionLastRow As Integer, lastColumn As Integer, currentColumn As Integer, lastRowForVesselId As Integer
   
    filter = "CSV files (*.csv),*.csv"
    MsgBox "Please select the Biolector raw data file", vbOKOnly
    rawDataFilename = Application.GetOpenFilename(filter)
    Set rawDataWorkbook = Application.Workbooks.Open(rawDataFilename)
    Set rawDataSheet = rawDataWorkbook.Worksheets(1)
    Set targetWorkbook = Application.ThisWorkbook
    Set targetSheet = targetWorkbook.Worksheets("Raw Data")
    Set mainSheet = targetWorkbook.Worksheets("Biolector Data")
   
    'Clear sheets and initializer header
    targetSheet.Cells.Clear
   mainSheet.Cells.Clear
    headersArray = Array("Well", "Time[h]", "Cali.pH", "pH", "Cali.DO", "DO", "Biomass", "MF Volume (A)", "MF Volume (B)", "Volumes", "Temperature", "Temperature Down [C]", "Temperature Water [C]", "O2 [%]", "CO2 [%]", "Shaker [rpm]", "Humidity [%rH]", "Time [h]")
    mainSheet.Range("A1:R1").Value = headersArray
       
    'Copy Biolector data to target sheet
    rawDataSheet.Cells.Copy Destination:=targetSheet.Cells
   
    Set patternFound = targetSheet.Range("A1:A" & Rows.Count).Find(What:="=====data=====")
    patternFoundRow = patternFound.Row
   
    'Define raw data section
    targetSheet.Activate
    rawDataSectionFirstRow = targetSheet.Range("A" & patternFoundRow).Offset(1, 0).Row
    rawDataSectionLastRow = targetSheet.Range("A" & patternFoundRow).Offset(1, 0).End(xlDown).Row
   
    'Delimit raw data section by semicolon
    targetSheet.Range("A" & rawDataSectionFirstRow, "A" & rawDataSectionLastRow).TextToColumns DataType:=xlDelimited, Semicolon:=True
   
    'Transform data from horizontal format to vertical format
    lastColumn = targetSheet.Cells(rawDataSectionFirstRow, targetSheet.Columns.Count).End(xlToLeft).Column
   
    'Convert lastColumn to letter
    lastColumnChar = Split(Cells(1, lastColumn).Address, "$")(1)
   
    'Paste array into vertical format, make upper bound +1 to make room for the headers
    timeArray = targetSheet.Range("G" & rawDataSectionFirstRow & ":" & lastColumnChar & rawDataSectionFirstRow).Value
    mainSheet.Range("B2:B" & WorksheetFunction.CountA(timeArray) + 1).Value = WorksheetFunction.Transpose(timeArray)
   
    'Loop over each vessel 8 times to collect raw data into arrays
    currentColumn = 2
   
    For i = 0 To 7
        currentColumn = currentColumn + 1
        currentColumnChar = Split(Cells(1, currentColumn).Address, "$")(1)
   
        dataArray = targetSheet.Range("G" & rawDataSectionFirstRow + (1 + (32 * i)) & ":" & lastColumnChar & rawDataSectionFirstRow + (1 + (32 * i))).Value
        mainSheet.Range(currentColumnChar & "2:" & currentColumnChar & WorksheetFunction.CountA(dataArray) + 1).Value = WorksheetFunction.Transpose(dataArray)
        ReDim dataArray(1 To lastColumn)
       
        'Fill down Well ID
        vesselID = targetSheet.Range("B" & rawDataSectionFirstRow + (1 + (32 * i))).Value
        lastRowForVesselId = mainSheet.Cells(Rows.Count, 2).End(xlUp).Row
        mainSheet.Range("A2:A" & lastRowForVesselId).Value = vesselID
    Next
 
End Sub
 
Sub Run_Biolector_Macro()
    Call parseBiolectorData
End Sub
