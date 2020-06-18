' Global constants
Private Const ScriptsFolder As String = "\Scripts\"
Private Const FilesFolder As String = "\Files\"

Sub GetUsernames()
    'Variable declaration
    Dim exportedUsersFileName As String
    Dim rangeToExport As Range
    Dim scriptPathAndName As String
    Dim headerArray As Variant
    Dim rangeArray As Variant
    
    'Variable initilization
    exportedUsersFileName = "ExpectedDisplayNames-" & (Environ$("Username")) & ".csv"
    rangeArray = Array(Range("K7:K16"))
    scriptPathAndName = "& '" & ThisWorkbook.path & ScriptsFolder & "getUsernames.ps1'"
    headerArray = Array("DisplayName")
    
    'methods to export data, run script and import processed data
    ExportDataToCSV exportedUsersFileName, rangeArray, headerArray
    RunScript scriptPathAndName
    ImportUsernamesFromCSV
End Sub

Sub ConfigADUsers()
    'variable declaration
    Dim exportedUsernamesFileName As String
    Dim userSets As Integer
    Dim headerArray As Variant
    Dim index As Integer
    Dim scriptPathAndName As String
    
    'variable initialization
    exportedUsernamesFileName = "UserData-" & (Environ$("Username")) & ".csv"
    scriptPathAndName = "& '" & ThisWorkbook.path & ScriptsFolder & "configADUsers.ps1'"
    
    'userSets stores the number of user sets to be handled (counts the number of cols with content starting from N7)
    userSets = ThisWorkbook.Sheets("Main").Range("N7", ThisWorkbook.Sheets("Main").Range("N7").End(xlToRight)).Columns.Count - 1
    
    'header names for each userSet
    headerArray = Array("Username", "OU")
    
    'Loop through the columns creating the users
    For index = 0 To (userSets - 1)
        ExportDataToCSV exportedUsernamesFileName, GetExportRanges(index), headerArray
        RunScript scriptPathAndName
    Next
    
    'Delete userName lists and show completed message
    ThisWorkbook.Sheets("Main").Range("O7:V16").ClearContents
    MsgBox ("Configuration of NH completed")
End Sub

Function GetExportRanges(index As Integer)
    'variable declaration
    Dim currentCol As Integer
    Dim firstRow As Integer
    Dim lastRow As Integer
    Dim accountRow As Integer
    Dim specificRow As Integer
    Dim positionRow As Integer
    Dim locationRow As Integer
    Dim usernamesRange As Range
    Dim textToSearch As String
    Dim ouRange As Range
    Dim ouRangeRow As Integer
    Dim ouRangeCol As Integer
    
    'Col/row constants
    currentCol = 15 + index 'number representation for Col O + index from CreateActiveDirectoryUsers loop
    firstRow = 7 'starting row for each set of users
    lastRow = 16 'maximum row for each set of users
    accountRow = 4 'row for account values
    specificRow = 18 'row for specific values
    positionRow = 5 'row for position values
    locationRow = 6 'row for location values
    ouRangeCol = 6 'OU col in Data sheet
    
    'usersNames range to save
    Set usernamesRange = Range(Cells(firstRow, currentCol), Cells(lastRow, currentCol))
        
    'concatenation of account, specific, position and location for each userSet
    textToSearch = Cells(accountRow, currentCol) & Cells(specificRow, currentCol) & Cells(positionRow, currentCol) _
        & Cells(locationRow, currentCol)
    'OU row depending on textToSearch
    ouRangeRow = Application.WorksheetFunction.Match(textToSearch, ThisWorkbook.Sheets("Data").Range("A1:A100"), 0)
        
    With ThisWorkbook.Sheets("Data")
        Set ouRange = .Range(.Cells(ouRangeRow, ouRangeCol), .Cells(ouRangeRow, ouRangeCol))
    End With
        
    GetExportRanges = Array(usernamesRange, ouRange)
End Function

Sub ImportUsernamesFromCSV()
    'Opens the CSV file created by PowerShell and copy the names to the first available column starting from O
    'Variable declaration
    Dim myWB As Workbook
    Dim csvFileName As String
    Dim initialRow As Integer
    Dim lastCol As Integer
    
    'Variable initialization
    Set myWB = ThisWorkbook
    csvFileName = "FoundUsernames-" & (Environ$("Username")) & ".csv"
    initialRow = 7 'Row where the usernames start until row 16
    
    'Opens csv file and activate the only existing sheet
    Workbooks.OpenText fileName:=myWB.path & FilesFolder & csvFileName, Local:=True
    Workbooks(csvFileName).Worksheets(1).Activate
    
    'Copy usernames and paste them on the first available column starting from O7
    Range("A1:A10").Copy
    
    myWB.Worksheets("Main").Activate
    
    lastCol = Cells(initialRow, Columns.Count).End(xlToLeft).Column
    Cells(initialRow, lastCol + 1).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
    
    Workbooks(csvFileName).Close
    
End Sub

Sub RunScript(script As String)
    'Launch powershell specified script and wait for it to finish
    'Variable declaration and initialization
    Dim wsh As Object
    Set wsh = VBA.CreateObject("WScript.Shell")
    Dim waitOnReturn As Boolean: waitOnReturn = True
    Dim windowStyle As Integer: windowStyle = 1
    
    'Run script using the created Shell object that will allow this to wait for the script to finish
    wsh.Run "powershell.exe " & script, windowStyle, waitOnReturn
    
End Sub

Sub ExportDataToCSV(fileName As String, rangeArray As Variant, headerArray As Variant)
    'Exports data located at rngToSave to csv fileName
    'variable declaration
    Dim myCSVFileName As String
    Dim myWB As Workbook
    Dim tempWB As Workbook
    Dim loopIndex As Integer
    Dim headerRow As Integer
    Dim dataRow As Integer

    'Turn off display alerts to avoid getting prompted and having to click accept every time
    Application.DisplayAlerts = False
    On Error GoTo err

    'Variable initialization
    Set myWB = ThisWorkbook
    myCSVFileName = myWB.path & FilesFolder & fileName
    headerRow = 1
    dataRow = 2

    'Add a new sheet to a temporal workbook
    Set tempWB = Application.Workbooks.Add(1)
    loopIndex = 0
    
    With tempWB.Sheets(1)
        'Loop through ranges/headers copying and pasting ranges for each header
        For Each header In headerArray
            .Cells(headerRow, loopIndex + 1) = header
        
            rangeArray(loopIndex).Copy
            .Cells(dataRow, loopIndex + 1).PasteSpecial xlPasteValues
        
            loopIndex = loopIndex + 1
        Next
    End With
    
    'Save it on myCSVFileName
    tempWB.SaveAs fileName:=myCSVFileName, FileFormat:=xlCSV, CreateBackup:=False
    tempWB.Close
    
    'Turn on alerts either on error or at the end of execution
err:
    Application.DisplayAlerts = True
End Sub
