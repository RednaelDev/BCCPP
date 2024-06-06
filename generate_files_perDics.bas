Attribute VB_Name = "Module4"
Sub generate_files_perDisc()
    ' Variables to get the file locations
    Dim sh As Worksheet
    Dim FldrPicker As FileDialog
    Dim myPath As String
    Dim myFile As String

    ' Variables for the data iteration
    Dim searchString As String
    Dim w As Integer
    Dim i As Integer
    Dim latestModifiedDate As Date
    Dim fileModifiedDate As Date
    Dim latestFile As String
    Dim fso As Object

    ' Arrays for the data iteration
    Dim Discip() As Variant, DiscipIndicator() As Variant
    
    Application.ScreenUpdating = False
    
    Discip() = Array("Mechanical", "Electrical", "Instrument")
    DiscipIndicator() = Array("(M)", "(E)", "(I)")

    ' Select the folder containing the files
    Set FldrPicker = Application.FileDialog(msoFileDialogFolderPicker)
    With FldrPicker
        .Title = "Select Folder"
        .AllowMultiSelect = False
        .ButtonName = "Confirm"
        If .Show = -1 Then
            myPath = .SelectedItems(1) & "\"
        Else
            End
        End If
    End With

    ' Iterate per worksheet and iterate per item
    w = 0
    Do While (w < 3) ' to iterate per sheet/discipline
        Set sh = ThisWorkbook.Sheets(Discip(w))
        sh.Activate

        Set fso = CreateObject("Scripting.FileSystemObject")
        i = 3
        Do While sh.Cells(i, 3) <> "" ' checks if there are more packages to be processed
            myFile = Dir(myPath)
            searchString = sh.Cells(i, 3).Value
            latestModifiedDate = 0
            latestFile = ""

            Do While myFile <> "" ' iterates through every file in the folder
                If (InStr(myFile, searchString) > 0) Then
                    If (InStr(myFile, DiscipIndicator(w)) > 0) Then
                        fileModifiedDate = fso.GetFile(myPath & myFile).DateLastModified
                        If fileModifiedDate > latestModifiedDate Then
                            latestModifiedDate = fileModifiedDate
                            latestFile = myPath & myFile
                        End If
                    End If
                End If
                myFile = Dir
            Loop
            
            If latestFile <> "" Then
                sh.Hyperlinks.Add Anchor:=sh.Cells(i, 14), Address:=latestFile, TextToDisplay:=" " & latestModifiedDate
            End If
            
            i = i + 1
        Loop
        w = w + 1
    Loop

MsgBox "Process completed!"

End Sub

