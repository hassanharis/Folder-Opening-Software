Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
Application.EnableEvents = False
' If Target.Address(False, False) = "B2" Then Call OpenAndPrintFailedFolders
Application.EnableEvents = True
End Sub

Private Sub CommandButton2_Click()
OpenFolderinCentralFiles
End Sub

Function FolderExists(folderPath As String) As Boolean
    On Error Resume Next
    FolderExists = (GetAttr(folderPath) And vbDirectory) = vbDirectory
    On Error GoTo 0
End Function

Sub OpenStaffBookings()

    Dim folderNamePart As String
    Dim folderPath As String
    Dim cellValue As String
    Dim foundFolder As Boolean
    Dim folder As String
    Dim folderPath2 As String
    Dim foundFolder2 As Boolean
    Dim folder2 As String
    
    ' Assuming the cell with the folder name part is in Sheet1, cell A1
    cellValue = Sheets("Training").Range("B3").Value

    ' Specify the base folder paths where you want to search for the folder
    ' Update these paths based on your requirements
    Dim basePath1 As String

    basePath1 = "R:\Central Files\Training Information\Clients Files\1. On Line\"


    ' Loop through the folders in the first base path

    folder = Dir(basePath1, vbDirectory)

    Do While folder <> ""
        ' Check if the folder name contains the specified value
        If InStr(1, folder, cellValue, vbTextCompare) > 0 Then
            ' Combine the base path and folder name to get the complete folder path
            folderPath = basePath1 & folder & "\"
            foundFolder = True
            'MsgBox folderPath, vbExclamation
            'Call Shell("explorer.exe """ & folderPath & """", vbNormalFocus)
            'MsgBox "Opened" & folder, vbExclamation
            'Exit Do
        End If
        folder = Dir
    Loop

    ' Check if a matching folder was found
    If foundFolder Then
        ' Open the folder using ShellExecute
        'Call Shell("explorer.exe """ & folderPath & """", vbNormalFocus)
        folder2 = Dir(folderPath, vbDirectory)

        Do While folder2 <> ""
            ' Check if the folder name contains the specified value
            If InStr(1, folder2, Year(Date) & " Staff Bookings", vbTextCompare) > 0 Then
                ' Combine the base path and folder name to get the complete folder path
                folderPath2 = folderPath & folder2
                foundFolder2 = True
                'MsgBox folderPath, vbExclamation
                Call Shell("explorer.exe """ & folderPath2 & """", vbNormalFocus)
                'MsgBox "Opened" & folder, vbExclamation
                'Exit Do
            End If
            folder2 = Dir
        Loop
    Else
        MsgBox "No matching folder found for: " & cellValue, vbExclamation
    End If
End Sub

Sub OpenClientFolderOfTraining()

    Dim folderNamePart As String
    Dim folderPath As String
    Dim cellValue As String

    ' Assuming the cell with the folder name part is in Sheet1, cell A1
    cellValue = Sheets("Training").Range("B3").Value

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

Sub OpenAndPrintFailedFolders()
    Dim folderPaths As Variant
    Dim pathArray As Variant
    Dim path As Variant
    Dim failedFolders As String
    Dim cellValue As String
    Dim folderPath As String
    Dim folderKeyword As String
    ' Get folder paths from cell B2
    folderPaths = Range("B2").Value
    
    ' Check if folderPaths is a string
    If VarType(folderPaths) = vbString Then
        ' Initialize the failedFolders string
        failedFolders = ""
        
        ' Split the string into an array using either a new line or a comma as the delimiter
        ' pathArray = Split(folderPaths, vbCrLf)

        folderPaths = Replace(folderPaths, vbCrLf, ",")
        folderPaths = Replace(folderPaths, vbLf, ",")
        
        pathArray = Split(folderPaths, ",")
        ' Loop through each path in the array
        For Each path In pathArray
            ' Trim leading and trailing spaces
            path = Trim(path)
            cellValue = path
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
                    MsgBox "Invalid first digit for determining the folder path." & firstDigit & folderPath, vbExclamation
                    failedFolders = failedFolders & path & vbCrLf
                    Exit Sub
            End Select
            
            path = folderPath
        
        
                        ' Check if the path is a valid directory
            If Len(Dir(path, vbDirectory)) > 0 Then
                ' Attempt to open the folder
                On Error Resume Next
                Call Shell("explorer.exe """ & path & """", vbNormalFocus)
                
                ' Check if an error occurred (folder not opened)
                If Err.Number <> 0 Then
                    ' Append the failed folder to the failedFolders string
                    failedFolders = failedFolders & cellValue & vbCrLf
                    ' Reset the error
                    Err.Clear
                End If
                On Error GoTo 0
            Else
                ' Append the invalid folder to the failedFolders string
                failedFolders = failedFolders & cellValue & " (Invalid Path)" & vbCrLf
            End If
        Next path
        
        ' Print the failed folders to cell D2
        Range("D2").Value = failedFolders
    Else
        MsgBox "Invalid input in cell B2. Please enter folder paths as a string."
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
    cellValue = Range("B2").Value

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
