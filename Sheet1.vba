Option Explicit
Private Sub Worksheet_Change(ByVal Target As Range)
'Application.EnableEvents = False
'If Target.Address(False, False) = "B2" Then Call OpenFolderinCentralFiles
'If Target.Address(False, False) = "B4" Then Call OpenQPFolderinPendingSites
'If Target.Address(False, False) = "B6" Then Call OpenClientFolderOfTraining
'If Target.Address(False, False) = "B8" Then Call SearchAntennas
'Application.EnableEvents = True
End Sub

Private Sub CommandButton2_Click()
OpenFolderinCentralFiles
End Sub

Function FolderExists(folderPath As String) As Boolean
    On Error Resume Next
    FolderExists = (GetAttr(folderPath) And vbDirectory) = vbDirectory
    On Error GoTo 0
End Function

Sub OpenClientFolderOfTraining()

    Dim folderNamePart As String
    Dim folderPath As String
    Dim cellValue As String

    ' Assuming the cell with the folder name part is in Sheet1, cell A1
    cellValue = Sheets("Sheet1").Range("B6").Value

    ' Specify the base folder paths where you want to search for the folder
    ' Update these paths based on your requirements
    Dim basePath1 As String

    basePath1 = "R:\Central Files\Training Information\Clients Files\1. On Line\"


    ' Loop through the folders in the first base path
    Dim foundFolder As Boolean
    Dim folder As String
    folder = Dir(basePath1, vbDirectory)

    Do While folder <> ""
        ' Check if the folder name contains the specified value
        If InStr(1, folder, cellValue, vbTextCompare) > 0 Then
            ' Combine the base path and folder name to get the complete folder path
            folderPath = basePath1 & folder
            foundFolder = True
            'MsgBox folderPath, vbExclamation
            Call Shell("explorer.exe """ & folderPath & """", vbNormalFocus)
            'MsgBox "Opened" & folder, vbExclamation
            'Exit Do
        End If
        folder = Dir
    Loop

    ' Check if a matching folder was found
    If foundFolder Then
        ' Open the folder using ShellExecute
        'Call Shell("explorer.exe """ & folderPath & """", vbNormalFocus)
        
    Else
        MsgBox "No matching folder found for: " & cellValue, vbExclamation
    End If
End Sub


Sub OpenQPFolderinPendingSites()

    Dim folderNamePart As String
    Dim folderPath As String
    Dim cellValue As String

    ' Assuming the cell with the folder name part is in Sheet1, cell A1
    cellValue = Sheets("Sheet1").Range("B4").Value

    ' Specify the base folder paths where you want to search for the folder
    ' Update these paths based on your requirements
    Dim basePath1 As String
    Dim basePath2 As String
    basePath1 = "R:\Central Files\Pending Sites\"
    basePath2 = "R:\Central Files\Pending Sites\SSMC TCI RFQ\"

    ' Loop through the folders in the first base path
    Dim foundFolder As Boolean
    Dim folder As String
    folder = Dir(basePath1, vbDirectory)

    Do While folder <> ""
        ' Check if the folder name contains the specified value
        If InStr(1, folder, cellValue, vbTextCompare) > 0 Then
            ' Combine the base path and folder name to get the complete folder path
            folderPath = basePath1 & folder
            foundFolder = True
            Call Shell("explorer.exe """ & folderPath & """", vbNormalFocus)
        End If
        folder = Dir
    Loop

    ' If the folder is not found in the first path, check the second path
    If Not foundFolder Then
        folder = Dir(basePath2, vbDirectory)
        Do While folder <> ""
            ' Check if the folder name contains the specified value
            If InStr(1, folder, cellValue, vbTextCompare) > 0 Then
                ' Combine the base path and folder name to get the complete folder path
                folderPath = basePath2 & folder
                foundFolder = True
                Call Shell("explorer.exe """ & folderPath & """", vbNormalFocus)
            End If
            folder = Dir
        Loop
    End If

    ' Check if a matching folder was found
    If foundFolder Then
        ' Open the folder using ShellExecute
       ' Call Shell("explorer.exe """ & folderPath & """", vbNormalFocus)
    Else
        MsgBox "No matching folder found for: " & cellValue, vbExclamation
    End If
End Sub





Sub OpenFolderinCentralFiles()
' Author Haris Hassan
' hharis11@hotmail.com
    
    Dim cellValue As String
    Dim folderKeyword As String
    Dim searchDirectory As String
    Dim folderPath As String
    ' Get the cell value from cell A1
    cellValue = Trim(Range("B2").Value)

    ' Check if the cell value is not empty
    If cellValue <> "" Then
        ' Determine the folder path based on the first digit
        If True Then
            Dim firstDigit As String
            firstDigit = Left(cellValue, 1)

            ' Determine the folder path based on the first digit
            Select Case firstDigit
                Case "1"
                    folderPath = "R:\Central Files\10000 - 19999  ACT\" & cellValue
                Case "2"
                    folderPath = "R:\Central Files\20000 - 29999  NSW\" & cellValue
                Case "3"
                    folderPath = "R:\Central Files\30000 - 39999  VIC\" & cellValue
                    If cellValue = "30396" Then
                        folderPath = "R:\Central Files\30000 - 39999  VIC\30396 - IBC"
                    End If
                
                Case "4"
                    folderPath = "R:\Central Files\40000 - 49999 QLD\" & cellValue
                Case "5"
                    folderPath = "R:\Central Files\50000 - 59999  SA\" & cellValue
                Case "6"
                    folderPath = "R:\Central Files\60000 - 69999 WA\" & cellValue
                Case "7"
                    folderPath = "R:\Central Files\70000 - 79999  TAS\" & cellValue
                Case "8"
                    folderPath = "R:\Central Files\80000 - 89999 NT\" & cellValue
                Case "0"
                    Select Case CStr(Left(cellValue, 5))
                        Case "00500"
                            ' Extract the text after the dash in the cell value
                            folderKeyword = " " & Split(cellValue, "-")(1)
                            
                            Dim foundFolder As Boolean
                            Dim basePath As String
                            basePath = "R:\Central Files\00000 - 04999 Other Reports\00500 - NAD\"
                        
                            ' Loop through the folders in the first base path
                            
                            Dim folder As String
                            folder = Dir(basePath, vbDirectory)
                        
                            Do While folder <> ""
                                ' Check if the folder name contains the specified value
                                If InStr(1, folder, folderKeyword, vbTextCompare) > 0 Then
                                    ' Combine the base path and folder name to get the complete folder path
                                    folderPath = basePath & folder
                                    foundFolder = True
                                    Exit Do
                                End If
                                folder = Dir
                            Loop
                            ' Set the search directory
                            ' Check if a matching folder was found
                            If foundFolder Then
                               
                            Else
                                MsgBox "No matching folder found for: " & folderKeyword & folderPath, vbExclamation
                            End If
                        
                        Case "00150"
                            folderPath = "R:\Central Files\00000 - 04999 Other Reports\" & CStr(Left(cellValue, 5)) & "\" & cellValue & "\"
                        Case "01065"
                            folderPath = "R:\Central Files\00000 - 04999 Other Reports\01065 - Radman Sales"
                        Case Else
                            folderPath = "R:\Central Files\00000 - 04999 Other Reports\" & CStr(Left(cellValue, 5)) & "\"
                    End Select
                ' Add more cases as needed
                Case Else
                    MsgBox "Invalid first digit for determining the folder path.", vbExclamation
                    Exit Sub
            End Select
            
            If FolderExists(folderPath) Then
            ' Open the existing folder in Windows Explorer
                Call Shell("explorer.exe """ & folderPath & """", vbNormalFocus)
            Else
                ' Prompt the user to create the folder
                'If MsgBox("The folder does not exist. Do you want to create it?", vbQuestion + vbYesNo) = vbYes Then
                    ' Create the folder
                    MsgBox "This folder does not exist." & folderPath, vbExclamation
                    'MkDir folderPath
                    ' Open the newly created folder in Windows Explorer
                    'Call Shell("explorer.exe """ & folderPath & """", vbNormalFocus)
               ' End If
            End If
            
        Else
            MsgBox "The value in cell A1 does not start with a digit.", vbExclamation
        End If
    Else
        ' Display a message if the cell is empty
        MsgBox "Please enter a valid value in cell A1.", vbExclamation
    End If
End Sub

Sub SearchAntennas()
    Dim folderContainingText As String
    Dim searchPath As String
    Dim result As String
    Dim tempFilePath As String
    Dim tempFileNumber As Integer
    Dim ws As Worksheet
    Dim cell As Range

    ' Set the worksheet (change "Sheet1" to your actual sheet name)
    Set ws = ThisWorkbook.Sheets("Sheet1")

    ' Get the text from cell B4
    folderContainingText = ws.Range("B8").Value

    ' Check if the text is not empty
    If folderContainingText <> "" Then
        ' Set the search path (change "C:\" to the directory where you want to search)
        searchPath = "R:\Temp\Temp\Prox5_Antennas_Pattern\"

        ' Reset result
        result = ""

        ' Find all folders and files in the specified directory
        Call FindFoldersAndFilesInDirectory(searchPath, folderContainingText, result)

        ' Create a temporary text file
        tempFilePath = Environ$("TEMP") & "\" & "SearchResults.txt"
        tempFileNumber = FreeFile
        Open tempFilePath For Output As tempFileNumber
        Print #tempFileNumber, result
        Close tempFileNumber

        ' Open the temporary text file for the user
        Shell "notepad.exe " & tempFilePath, vbNormalFocus
    Else
        ' Display a message if the cell is empty
        MsgBox "Please enter a valid text in cell B4.", vbExclamation
    End If
End Sub

Sub FindFoldersAndFilesInDirectory(ByVal folderPath As String, ByVal searchText As String, ByRef result As String)
    Dim fso As Object
    Dim folder As Object
    Dim subfolder As Object
    Dim file As Object

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(folderPath)

    ' Loop through each subfolder and find folders and files with names containing the specified text
    For Each subfolder In folder.SubFolders
        If InStr(1, subfolder.Name, searchText, vbTextCompare) > 0 Then
            result = result & subfolder.path & vbCrLf
        End If
        Call FindFoldersAndFilesInDirectory(subfolder.path, searchText, result)
    Next subfolder

    ' Loop through each file and find files with names containing the specified text
    For Each file In folder.Files
        If InStr(1, file.Name, searchText, vbTextCompare) > 0 Then
            result = result & file.path & vbCrLf
        End If
    Next file
End Sub


Sub FindHistoryOfProject()
' Author Haris Hassan
' hharis11@hotmail.com
    Dim cellValue As String
    Dim folderKeyword As String
    Dim searchDirectory As String
    Dim subfolderList As String
    Dim folderPath As String
    Dim tempFilePath As String
    Dim tempFileNumber As Integer
    Dim ws As Worksheet

    ' Get the cell value from cell A1
    cellValue = Range("B6").Value

    ' Check if the cell value is not empty
    If cellValue <> "" Then
        ' Determine the folder path based on the first digit
            Dim firstDigit As Integer
            firstDigit = CInt(Left(cellValue, 1))

            ' Determine the folder path based on the first digit
            Select Case firstDigit
                Case 1
                    folderPath = "R:\Central Files\10000 - 19999  ACT\" & cellValue & "\"
                Case 2
                    folderPath = "R:\Central Files\20000 - 29999  NSW\" & cellValue & "\"
                Case 3
                    folderPath = "R:\Central Files\30000 - 39999  VIC\" & cellValue & "\"
                Case 4
                    folderPath = "R:\Central Files\40000 - 49999 QLD\" & cellValue & "\"
                Case 5
                    folderPath = "R:\Central Files\50000 - 59999  SA\" & cellValue & "\"
                Case 6
                    folderPath = "R:\Central Files\60000 - 69999 WA\" & cellValue & "\"
                Case 7
                    folderPath = "R:\Central Files\70000 - 79999  TAS\" & cellValue & "\"
                Case 8
                    folderPath = "R:\Central Files\80000 - 89999 NT\" & cellValue & "\"
                Case "0"
                    Select Case CStr(Left(cellValue, 5))
                        Case "00500"
                            ' Extract the text after the dash in the cell value
                            folderKeyword = " " & Split(cellValue, "-")(1)
                            
                            Dim foundFolder As Boolean
                            Dim basePath As String
                            basePath = "R:\Central Files\00000 - 04999 Other Reports\00500 - NAD\"
                        
                            ' Loop through the folders in the first base path
                            
                            Dim folder As String
                            folder = Dir(basePath, vbDirectory)
                        
                            Do While folder <> ""
                                ' Check if the folder name contains the specified value
                                If InStr(1, folder, folderKeyword, vbTextCompare) > 0 Then
                                    ' Combine the base path and folder name to get the complete folder path
                                    folderPath = basePath & folder
                                    foundFolder = True
                                    Exit Do
                                End If
                                folder = Dir
                            Loop
                            ' Set the search directory
                            ' Check if a matching folder was found
                            If foundFolder Then
                               
                            Else
                                MsgBox "No matching folder found for: " & folderKeyword & folderPath, vbExclamation
                            End If
                        
                        Case "00150"
                            folderPath = "R:\Central Files\00000 - 04999 Other Reports\" & CStr(Left(cellValue, 5)) & "\" & cellValue & "\"
                        Case "01065"
                            folderPath = "R:\Central Files\00000 - 04999 Other Reports\01065 - Radman Sales"
                        Case Else
                            folderPath = "R:\Central Files\00000 - 04999 Other Reports\" & CStr(Left(cellValue, 5)) & "\"
                    End Select
                ' Add more cases as needed
                Case Else
                    MsgBox "Invalid first digit for determining the folder path.", vbExclamation
                    Exit Sub
            End Select
            
                ' Check if the folder path is not empty
            If folderPath <> "" Then
                ' Reset subfolder list
                subfolderList = ""
        
                ' Find all immediate subfolders in the specified directory
                Call FindImmediateSubfolders(folderPath, subfolderList)

                ' Create a temporary text file
                tempFilePath = Environ$("TEMP") & "\" & "SubfolderList.txt"
                tempFileNumber = FreeFile
                Open tempFilePath For Output As tempFileNumber
                Print #tempFileNumber, subfolderList
                Close tempFileNumber
        
                ' Open the temporary text file for the user
                Shell "notepad.exe " & tempFilePath, vbNormalFocus
            Else
                ' Display a message if the cell is empty
                MsgBox "Please enter a valid folder path in cell B4.", vbExclamation
            End If

    Else
        ' Display a message if the cell is empty
        MsgBox "Please enter a valid value in cell A1.", vbExclamation
    End If
End Sub

Sub FindSubfolders(ByVal folderPath As String, ByVal subfolderList As String)
    Dim fso As Object
    Dim folder As Object
    Dim subfolder As Object

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(folderPath)

    ' Loop through each subfolder and add it to the subfolder list
    For Each subfolder In folder.SubFolders
        subfolderList = subfolderList & subfolder.path & vbCrLf
        Call FindSubfolders(subfolder.path, subfolderList) ' Recursively find subfolders
    Next subfolder
End Sub

Sub FindImmediateSubfolders(ByVal folderPath As String, ByVal subfolderList As String)
    Dim fso As Object
    Dim folder As Object
    Dim subfolder As Object

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(folderPath)

    ' Loop through each immediate subfolder and add it to the subfolder list
    For Each subfolder In folder.SubFolders
        subfolderList = subfolderList & subfolder.path & vbCrLf
    Next subfolder
End Sub

